using Microsoft.Graph;
using Microsoft.Graph.Models;
using TeamsChatCrawler.Models;

namespace TeamsChatCrawler.Services;

/// <summary>
/// Fetches messages (including hosted content / attachments) from a Teams chat
/// within a given time window.
/// Ref: https://learn.microsoft.com/en-us/graph/api/chat-list-messages
/// </summary>
public class MessageService(GraphServiceClient graph)
{
    private const int PageSize = 50;

    /// <summary>
    /// Returns all messages posted in <paramref name="chatId"/> between
    /// <see cref="CrawlOptions.WindowStart"/> and <see cref="CrawlOptions.WindowEnd"/>.
    /// </summary>
    public async Task<List<ChatMessage>> GetMessagesInWindowAsync(
        string chatId,
        string chatDisplayName,
        CrawlOptions options,
        CancellationToken ct = default)
    {
        var results = new List<ChatMessage>();

        // Graph supports $filter on createdDateTime for chat messages.
        // Format: createdDateTime ge {start} and createdDateTime le {end}
        var startIso = options.WindowStart.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
        var endIso   = options.WindowEnd.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
        var filter   = $"createdDateTime ge {startIso} and createdDateTime le {endIso}";

        var response = await graph.Me.Chats[chatId].Messages.GetAsync(req =>
        {
            req.QueryParameters.Top    = PageSize;
            req.QueryParameters.Filter = filter;
        }, ct);

        var pageIterator = PageIterator<ChatMessage, ChatMessageCollectionResponse>.CreatePageIterator(
            graph,
            response!,
            msg =>
            {
                results.Add(msg);
                return true;
            });

        await pageIterator.IterateAsync(ct);

        return results;
    }
}
