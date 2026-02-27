using System.Text.Json.Serialization;

namespace TeamsChatCrawler;

public sealed class GraphPage<T>
{
    [JsonPropertyName("value")]
    public List<T> Value { get; set; } = [];

    [JsonPropertyName("@odata.nextLink")]
    public string? NextLink { get; set; }
}

public sealed class ChatResource
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("topic")]
    public string? Topic { get; set; }

    [JsonPropertyName("chatType")]
    public string? ChatType { get; set; }
}

public sealed class ChatMessageResource
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("createdDateTime")]
    public DateTimeOffset? CreatedDateTime { get; set; }

    [JsonPropertyName("lastModifiedDateTime")]
    public DateTimeOffset? LastModifiedDateTime { get; set; }

    [JsonPropertyName("body")]
    public ItemBody? Body { get; set; }

    [JsonPropertyName("attachments")]
    public List<ChatMessageAttachment> Attachments { get; set; } = [];
}

public sealed class ItemBody
{
    [JsonPropertyName("contentType")]
    public string? ContentType { get; set; }

    [JsonPropertyName("content")]
    public string? Content { get; set; }
}

public sealed class ChatMessageAttachment
{
    [JsonPropertyName("id")]
    public string? Id { get; set; }

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("contentType")]
    public string? ContentType { get; set; }

    [JsonPropertyName("contentUrl")]
    public string? ContentUrl { get; set; }
}

public sealed class ChatGroupingCollection
{
    [JsonPropertyName("groups")]
    public List<ChatGrouping> Groups { get; set; } = [];
}

public sealed class ChatGrouping
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("chatIds")]
    public List<string> ChatIds { get; set; } = [];

    [JsonPropertyName("chatTitles")]
    public List<string> ChatTitles { get; set; } = [];
}

public sealed record CrawlSummary(int ChatsScanned, int MessagesExported, int AttachmentsDownloaded);
