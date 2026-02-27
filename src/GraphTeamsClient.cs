using System.Net.Http.Json;
using System.Text.Json;

namespace TeamsChatCrawler;

public sealed class GraphTeamsClient(HttpClient httpClient)
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    public async Task<IReadOnlyList<ChatResource>> GetMyChatsAsync(CancellationToken cancellationToken)
    {
        var chats = new List<ChatResource>();
        var endpoint = "me/chats?$top=50";

        while (!string.IsNullOrEmpty(endpoint))
        {
            var page = await GetAsync<GraphPage<ChatResource>>(endpoint, cancellationToken);
            chats.AddRange(page.Value);
            endpoint = NormalizeNextLink(page.NextLink);
        }

        return chats;
    }

    public async Task<IReadOnlyList<ChatMessageResource>> GetChatMessagesAsync(
        string chatId,
        DateTimeOffset fromInclusive,
        DateTimeOffset toInclusive,
        CancellationToken cancellationToken)
    {
        var messages = new List<ChatMessageResource>();
        var endpoint = $"chats/{Uri.EscapeDataString(chatId)}/messages?$top=50";

        while (!string.IsNullOrEmpty(endpoint))
        {
            var page = await GetAsync<GraphPage<ChatMessageResource>>(endpoint, cancellationToken);
            messages.AddRange(page.Value.Where(m => IsInsideWindow(m, fromInclusive, toInclusive)));
            endpoint = NormalizeNextLink(page.NextLink);
        }

        return messages;
    }

    public async Task DownloadAttachmentAsync(string contentUrl, string outputPath, CancellationToken cancellationToken)
    {
        using var response = await httpClient.GetAsync(contentUrl, cancellationToken);
        response.EnsureSuccessStatusCode();

        await using var source = await response.Content.ReadAsStreamAsync(cancellationToken);
        await using var target = File.Create(outputPath);
        await source.CopyToAsync(target, cancellationToken);
    }

    public async Task SaveJsonAsync<T>(T payload, string outputPath, CancellationToken cancellationToken)
    {
        await using var stream = File.Create(outputPath);
        await JsonSerializer.SerializeAsync(stream, payload, JsonOptions, cancellationToken);
    }

    private static bool IsInsideWindow(ChatMessageResource message, DateTimeOffset fromInclusive, DateTimeOffset toInclusive)
    {
        var created = message.CreatedDateTime ?? message.LastModifiedDateTime;
        if (created is null)
        {
            return false;
        }

        return created >= fromInclusive && created <= toInclusive;
    }

    private async Task<T> GetAsync<T>(string endpointOrAbsoluteUrl, CancellationToken cancellationToken)
    {
        using var response = await httpClient.GetAsync(endpointOrAbsoluteUrl, cancellationToken);
        response.EnsureSuccessStatusCode();

        var payload = await response.Content.ReadFromJsonAsync<T>(cancellationToken: cancellationToken);
        return payload ?? throw new InvalidOperationException($"Empty response from Graph endpoint {endpointOrAbsoluteUrl}");
    }

    private static string? NormalizeNextLink(string? nextLink)
    {
        if (string.IsNullOrWhiteSpace(nextLink))
        {
            return null;
        }

        const string graphPrefix = "https://graph.microsoft.com/v1.0/";
        return nextLink.StartsWith(graphPrefix, StringComparison.OrdinalIgnoreCase)
            ? nextLink[graphPrefix.Length..]
            : nextLink;
    }
}
