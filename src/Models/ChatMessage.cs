namespace TeamsChatCrawler.Models;

public class CrawledMessage
{
    public string MessageId { get; set; } = string.Empty;
    public string ChatId { get; set; } = string.Empty;
    public string ChatDisplayName { get; set; } = string.Empty;
    public string SenderName { get; set; } = string.Empty;
    public DateTimeOffset? CreatedDateTime { get; set; }
    public string Body { get; set; } = string.Empty;
    public List<CrawledAttachment> Attachments { get; set; } = new();
}

public class CrawledAttachment
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public string? ContentUrl { get; set; }
    public string? DownloadedFilePath { get; set; }
}
