using Microsoft.Graph;
using Microsoft.Graph.Models;
using TeamsChatCrawler.Models;

namespace TeamsChatCrawler.Services;

/// <summary>
/// Coordinates the full crawl: resolve chats → fetch messages → download attachments.
/// </summary>
public class CrawlerOrchestrator(
    ChatService chatService,
    MessageService messageService,
    DownloadService downloadService)
{
    public async Task RunAsync(CrawlOptions options, string outputRoot, CancellationToken ct = default)
    {
        Console.WriteLine();
        Console.WriteLine($"Crawl window : {options.WindowStart:yyyy-MM-dd} to {options.WindowEnd:yyyy-MM-dd}");
        Console.WriteLine($"Chat filters : {(options.ChatFilters.Count > 0 ? string.Join(", ", options.ChatFilters) : "(all chats)")}");
        Console.WriteLine();

        // 1. Resolve which chats to crawl
        Console.Write("Fetching chats...");
        var chats = await chatService.GetChatsAsync(options.ChatFilters, ct);
        Console.WriteLine($" {chats.Count} chat(s) selected.");

        if (chats.Count == 0)
        {
            Console.WriteLine("No chats matched the provided filters. Exiting.");
            return;
        }

        var allMessages = new List<CrawledMessage>();
        var totalAttachments = 0;

        // 2. Per chat: fetch messages in the time window
        for (var i = 0; i < chats.Count; i++)
        {
            var chat = chats[i];
            var displayName = ChatService.GetChatDisplayName(chat);

            Console.WriteLine();
            Console.WriteLine($"[{i + 1}/{chats.Count}] Chat: {displayName}");
            Console.Write("  Fetching messages...");

            var rawMessages = await messageService.GetMessagesInWindowAsync(
                chat.Id!, displayName, options, ct);

            Console.WriteLine($" {rawMessages.Count} message(s) in window.");

            // 3. Per message: map to domain model and download attachments
            foreach (var raw in rawMessages)
            {
                if (raw.MessageType != ChatMessageType.Message) continue; // skip system events

                var crawled = new CrawledMessage
                {
                    MessageId       = raw.Id ?? string.Empty,
                    ChatId          = chat.Id ?? string.Empty,
                    ChatDisplayName = displayName,
                    SenderName      = raw.From?.User?.DisplayName
                                   ?? raw.From?.Application?.DisplayName
                                   ?? "Unknown",
                    CreatedDateTime = raw.CreatedDateTime,
                    Body            = raw.Body?.Content ?? string.Empty
                };

                // Build per-chat / per-date output directory
                var safeChat  = SanitizeFolderName(displayName);
                var dateStr   = raw.CreatedDateTime?.UtcDateTime.ToString("yyyy-MM-dd") ?? "unknown-date";
                var attachDir = Path.Combine(outputRoot, safeChat, dateStr, "attachments");

                if (raw.Attachments?.Count > 0)
                {
                    Console.Write($"  Downloading {raw.Attachments.Count} attachment(s)...");
                    crawled.Attachments = await downloadService.DownloadAttachmentsAsync(
                        chat.Id!, raw, attachDir, ct);
                    Console.WriteLine(" done.");
                    totalAttachments += crawled.Attachments.Count;
                }

                allMessages.Add(crawled);
            }
        }

        // 4. Write a consolidated text summary
        await WriteSummaryAsync(allMessages, outputRoot, options, ct);

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Done. {allMessages.Count} messages and {totalAttachments} attachment(s) saved to: {outputRoot}");
        Console.ResetColor();
    }

    private static async Task WriteSummaryAsync(
        List<CrawledMessage> messages,
        string outputRoot,
        CrawlOptions options,
        CancellationToken ct)
    {
        Directory.CreateDirectory(outputRoot);
        var summaryPath = Path.Combine(outputRoot, "summary.txt");

        await using var sw = new StreamWriter(summaryPath, append: false);
        await sw.WriteLineAsync($"Teams Chat Crawl Summary");
        await sw.WriteLineAsync($"Reference date : {options.ReferenceDate}");
        await sw.WriteLineAsync($"Window         : {options.WindowStart:yyyy-MM-dd} to {options.WindowEnd:yyyy-MM-dd}");
        await sw.WriteLineAsync($"Total messages : {messages.Count}");
        await sw.WriteLineAsync(new string('-', 60));

        foreach (var msg in messages.OrderBy(m => m.CreatedDateTime))
        {
            await sw.WriteLineAsync();
            await sw.WriteLineAsync($"Chat    : {msg.ChatDisplayName}");
            await sw.WriteLineAsync($"From    : {msg.SenderName}");
            await sw.WriteLineAsync($"Date    : {msg.CreatedDateTime:yyyy-MM-dd HH:mm:ss} UTC");
            await sw.WriteLineAsync($"Message : {StripHtml(msg.Body)}");

            foreach (var att in msg.Attachments)
            {
                var local = att.DownloadedFilePath is not null
                    ? $" -> {att.DownloadedFilePath}"
                    : " (not downloaded)";
                await sw.WriteLineAsync($"  Attachment: {att.Name}{local}");
            }
        }

        Console.WriteLine($"  Summary written to: {summaryPath}");
    }

    private static string StripHtml(string html)
    {
        // Very lightweight HTML stripping — not a full parser, but sufficient for Teams message bodies
        if (string.IsNullOrEmpty(html)) return html;
        return System.Text.RegularExpressions.Regex.Replace(html, "<[^>]*>", " ")
               .Replace("&nbsp;", " ")
               .Replace("&amp;", "&")
               .Replace("&lt;", "<")
               .Replace("&gt;", ">")
               .Replace("&quot;", "\"")
               .Trim();
    }

    private static string SanitizeFolderName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Concat(name.Select(c => invalid.Contains(c) ? '_' : c))
                     .Trim()
                     .Replace(' ', '_');
    }
}
