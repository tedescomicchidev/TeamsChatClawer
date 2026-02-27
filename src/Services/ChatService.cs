using Microsoft.Graph;
using Microsoft.Graph.Models;
using TeamsChatCrawler.Models;

namespace TeamsChatCrawler.Services;

/// <summary>
/// Retrieves Teams chats from Microsoft Graph and applies optional name/topic filters.
/// Ref: https://learn.microsoft.com/en-us/graph/api/chat-list
/// </summary>
public class ChatService(GraphServiceClient graph)
{
    private const int PageSize = 50;

    /// <summary>
    /// Returns all chats the signed-in user is a member of, optionally filtered
    /// by a set of display-name / topic substrings (case-insensitive OR match).
    /// </summary>
    public async Task<List<Chat>> GetChatsAsync(List<string> filters, CancellationToken ct = default)
    {
        var allChats = new List<Chat>();

        var response = await graph.Me.Chats.GetAsync(req =>
        {
            req.QueryParameters.Top = PageSize;
            req.QueryParameters.Expand = ["members"];
        }, ct);

        // Use the SDK's PageIterator to handle OData @nextLink pages automatically
        var pageIterator = PageIterator<Chat, ChatCollectionResponse>.CreatePageIterator(
            graph,
            response!,
            chat =>
            {
                allChats.Add(chat);
                return true; // continue iterating
            });

        await pageIterator.IterateAsync(ct);

        if (filters.Count == 0)
            return allChats;

        return allChats.Where(c => MatchesAnyFilter(c, filters)).ToList();
    }

    /// <summary>
    /// Resolves a human-readable display name for a chat.
    /// - OneOnOne: "User A / User B"
    /// - Group: uses Topic if set, else lists member names
    /// - Meeting: uses Topic
    /// </summary>
    public static string GetChatDisplayName(Chat chat)
    {
        if (!string.IsNullOrWhiteSpace(chat.Topic))
            return chat.Topic;

        var members = chat.Members?
            .OfType<AadUserConversationMember>()
            .Select(m => m.DisplayName ?? "Unknown")
            .ToList() ?? [];

        return members.Count > 0 ? string.Join(" / ", members) : chat.Id ?? "Unknown Chat";
    }

    private static bool MatchesAnyFilter(Chat chat, List<string> filters)
    {
        var topic = chat.Topic ?? string.Empty;

        var memberNames = chat.Members?
            .OfType<AadUserConversationMember>()
            .Select(m => m.DisplayName ?? string.Empty) ?? [];

        var searchable = string.Join(" ", new[] { topic }.Concat(memberNames))
                              .ToLowerInvariant();

        return filters.Any(f => searchable.Contains(f.ToLowerInvariant()));
    }
}
