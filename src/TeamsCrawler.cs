using System.Text.Json;
using System.Text.RegularExpressions;

namespace TeamsChatCrawler;

public sealed class TeamsCrawler(GraphTeamsClient graphClient)
{
    private static readonly Regex UnsafePathCharsRegex = new("[^a-zA-Z0-9._-]+", RegexOptions.Compiled);

    public async Task<CrawlSummary> RunAsync(CliOptions options, string outputRoot, CancellationToken cancellationToken)
    {
        var allChats = await graphClient.GetMyChatsAsync(cancellationToken);
        var matchingChats = await FilterChatsAsync(allChats, options, cancellationToken);

        var toInclusive = options.AnchorDate.ToDateTime(TimeOnly.MaxValue, DateTimeKind.Utc);
        var fromInclusive = options.AnchorDate.AddDays(-5).ToDateTime(TimeOnly.MinValue, DateTimeKind.Utc);

        var messagesExported = 0;
        var attachmentsDownloaded = 0;

        foreach (var chat in matchingChats)
        {
            var chatFolder = Path.Combine(outputRoot, Sanitize(chat.Topic ?? chat.Id));
            Directory.CreateDirectory(chatFolder);

            var messages = await graphClient.GetChatMessagesAsync(chat.Id, fromInclusive, toInclusive, cancellationToken);
            messagesExported += messages.Count;

            var messagesPath = Path.Combine(chatFolder, "messages.json");
            await graphClient.SaveJsonAsync(messages, messagesPath, cancellationToken);

            var attachmentFolder = Path.Combine(chatFolder, "attachments");
            Directory.CreateDirectory(attachmentFolder);

            foreach (var message in messages)
            {
                foreach (var attachment in message.Attachments.Where(ShouldDownloadAttachment))
                {
                    var fileName = BuildAttachmentFileName(message.Id, attachment);
                    var attachmentPath = Path.Combine(attachmentFolder, fileName);
                    await graphClient.DownloadAttachmentAsync(attachment.ContentUrl!, attachmentPath, cancellationToken);
                    attachmentsDownloaded++;
                }
            }
        }

        return new CrawlSummary(matchingChats.Count, messagesExported, attachmentsDownloaded);
    }

    private static bool ShouldDownloadAttachment(ChatMessageAttachment attachment)
    {
        if (string.IsNullOrWhiteSpace(attachment.ContentUrl))
        {
            return false;
        }

        var contentType = attachment.ContentType ?? string.Empty;
        return contentType.Contains("image", StringComparison.OrdinalIgnoreCase)
               || contentType.Contains("application/", StringComparison.OrdinalIgnoreCase)
               || contentType.Contains("reference", StringComparison.OrdinalIgnoreCase);
    }

    private async Task<List<ChatResource>> FilterChatsAsync(
        IReadOnlyList<ChatResource> allChats,
        CliOptions options,
        CancellationToken cancellationToken)
    {
        if (options.ChatFilters.Count == 0 && options.GroupFilters.Count == 0)
        {
            return allChats.ToList();
        }

        var includedChatIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var normalizedChatFilters = options.ChatFilters.Select(static c => c.Trim()).Where(static c => !string.IsNullOrWhiteSpace(c)).ToList();

        foreach (var chat in allChats)
        {
            var title = chat.Topic ?? string.Empty;
            if (normalizedChatFilters.Any(filter => title.Contains(filter, StringComparison.OrdinalIgnoreCase)))
            {
                includedChatIds.Add(chat.Id);
            }
        }

        if (options.GroupFilters.Count > 0)
        {
            var groupings = await LoadGroupingsAsync(options, cancellationToken);
            var requested = new HashSet<string>(options.GroupFilters, StringComparer.OrdinalIgnoreCase);

            foreach (var grouping in groupings.Where(g => requested.Contains(g.Name)))
            {
                foreach (var chatId in grouping.ChatIds)
                {
                    includedChatIds.Add(chatId);
                }

                foreach (var chat in allChats)
                {
                    var title = chat.Topic ?? string.Empty;
                    if (grouping.ChatTitles.Any(filter => title.Contains(filter, StringComparison.OrdinalIgnoreCase)))
                    {
                        includedChatIds.Add(chat.Id);
                    }
                }
            }
        }

        return allChats.Where(chat => includedChatIds.Contains(chat.Id)).ToList();
    }

    private static async Task<IReadOnlyList<ChatGrouping>> LoadGroupingsAsync(CliOptions options, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(options.GroupConfigPath))
        {
            throw new InvalidOperationException("--group was provided but no --group-config JSON file was provided.");
        }

        var path = Path.GetFullPath(options.GroupConfigPath);
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("Group configuration file not found.", path);
        }

        await using var stream = File.OpenRead(path);
        var payload = await JsonSerializer.DeserializeAsync<ChatGroupingCollection>(stream, cancellationToken: cancellationToken);
        return payload?.Groups ?? [];
    }

    private static string BuildAttachmentFileName(string messageId, ChatMessageAttachment attachment)
    {
        var baseName = attachment.Name;
        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = attachment.Id ?? "attachment";
        }

        return $"{Sanitize(messageId)}_{Sanitize(baseName)}";
    }

    private static string Sanitize(string input)
    {
        var sanitized = UnsafePathCharsRegex.Replace(input, "_").Trim('_');
        return string.IsNullOrWhiteSpace(sanitized) ? "untitled" : sanitized;
    }
}
