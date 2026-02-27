using Microsoft.Graph;
using Microsoft.Graph.Models;
using TeamsChatCrawler.Models;

namespace TeamsChatCrawler.Services;

/// <summary>
/// Downloads file and image attachments from Teams chat messages.
///
/// Teams exposes two kinds of binary content:
/// 1. Hosted content (inline images) via the chatMessage/hostedContents endpoint.
///    Ref: https://learn.microsoft.com/en-us/graph/api/chatmessagehostedcontent-get
/// 2. File attachments backed by SharePoint/OneDrive – the contentUrl points to a
///    SharePoint driveItem download URL which can be fetched directly.
///    Ref: https://learn.microsoft.com/en-us/graph/api/driveitem-get-content
/// </summary>
public class DownloadService(GraphServiceClient graph, HttpClient httpClient)
{
    private static readonly HashSet<string> ImageMimeTypes =
        ["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp", "image/svg+xml"];

    private static readonly HashSet<string> FileMimeTypes =
        [
            "application/pdf",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "application/msword",
            "application/vnd.ms-excel",
            "application/vnd.ms-powerpoint",
            "text/plain",
            "application/zip",
            "reference"  // Teams uses "reference" for SharePoint file cards
        ];

    /// <summary>
    /// Downloads all attachments (images + documents) from a single message.
    /// Returns CrawledAttachment records with the local file path filled in.
    /// </summary>
    public async Task<List<CrawledAttachment>> DownloadAttachmentsAsync(
        string chatId,
        ChatMessage message,
        string outputDir,
        CancellationToken ct = default)
    {
        var downloaded = new List<CrawledAttachment>();

        if (message.Attachments is null || message.Attachments.Count == 0)
            return downloaded;

        Directory.CreateDirectory(outputDir);

        foreach (var attachment in message.Attachments)
        {
            if (attachment.Id is null) continue;

            var crawled = new CrawledAttachment
            {
                Id          = attachment.Id,
                Name        = attachment.Name ?? attachment.Id,
                ContentType = attachment.ContentType ?? string.Empty,
                ContentUrl  = attachment.ContentUrl
            };

            try
            {
                // Case 1: inline / hosted content (e.g. pasted images)
                if (attachment.ContentType == "application/vnd.microsoft.teams.file.download.info"
                    || string.IsNullOrEmpty(attachment.ContentUrl))
                {
                    var path = await DownloadHostedContentAsync(
                        chatId, message.Id!, attachment.Id, crawled.Name, outputDir, ct);
                    crawled.DownloadedFilePath = path;
                }
                // Case 2: SharePoint / reference file attachment
                else if (IsDownloadableAttachment(attachment.ContentType))
                {
                    var path = await DownloadUrlContentAsync(
                        attachment.ContentUrl!, crawled.Name, outputDir, ct);
                    crawled.DownloadedFilePath = path;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"  [WARN] Could not download attachment '{crawled.Name}': {ex.Message}");
            }

            downloaded.Add(crawled);
        }

        return downloaded;
    }

    /// <summary>
    /// Downloads hosted content (inline images) via the Graph hostedContents endpoint.
    /// </summary>
    private async Task<string?> DownloadHostedContentAsync(
        string chatId,
        string messageId,
        string contentId,
        string name,
        string outputDir,
        CancellationToken ct)
    {
        // GET /me/chats/{chatId}/messages/{messageId}/hostedContents/{id}/$value
        var stream = await graph.Me.Chats[chatId]
            .Messages[messageId]
            .HostedContents[contentId]
            .Content
            .GetAsync(cancellationToken: ct);

        if (stream is null) return null;

        var filePath = BuildSafeFilePath(outputDir, name);
        await using var fs = File.OpenWrite(filePath);
        await stream.CopyToAsync(fs, ct);
        return filePath;
    }

    /// <summary>
    /// Downloads a file via a direct URL (SharePoint pre-authenticated download URL).
    /// The Graph SDK's HttpClient already carries the Bearer token, so we use it.
    /// </summary>
    private async Task<string?> DownloadUrlContentAsync(
        string url,
        string name,
        string outputDir,
        CancellationToken ct)
    {
        using var response = await httpClient.GetAsync(url, ct);
        response.EnsureSuccessStatusCode();

        var filePath = BuildSafeFilePath(outputDir, name);
        await using var fs = File.OpenWrite(filePath);
        await response.Content.CopyToAsync(fs, ct);
        return filePath;
    }

    private static bool IsDownloadableAttachment(string? contentType)
    {
        if (string.IsNullOrEmpty(contentType)) return false;
        // Match exact types or prefix-match image/* etc.
        return FileMimeTypes.Contains(contentType)
            || ImageMimeTypes.Contains(contentType)
            || contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildSafeFilePath(string dir, string name)
    {
        // Strip illegal path characters from the file name
        var safeName = string.Concat(name.Split(Path.GetInvalidFileNameChars()));
        if (string.IsNullOrWhiteSpace(safeName)) safeName = "attachment";

        var filePath = Path.Combine(dir, safeName);

        // Handle duplicates: file.txt -> file_1.txt -> file_2.txt
        if (!File.Exists(filePath)) return filePath;

        var ext  = Path.GetExtension(safeName);
        var stem = Path.GetFileNameWithoutExtension(safeName);
        var i    = 1;
        do
        {
            filePath = Path.Combine(dir, $"{stem}_{i++}{ext}");
        } while (File.Exists(filePath));

        return filePath;
    }
}
