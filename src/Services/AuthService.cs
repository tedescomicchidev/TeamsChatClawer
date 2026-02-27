using Azure.Identity;
using Microsoft.Graph;

namespace TeamsChatCrawler.Services;

/// <summary>
/// Provides an authenticated Microsoft Graph client using the device code flow.
/// Users sign in interactively via the terminal — no client secret required.
/// </summary>
public static class AuthService
{
    // Scopes required by the Teams Chat REST APIs:
    //   Chat.Read          – list chats and read messages
    //   Chat.ReadBasic     – list chats (fallback for non-member chats)
    //   Files.Read.All     – download SharePoint/OneDrive file attachments
    //   User.Read          – resolve the signed-in user's profile
    private static readonly string[] Scopes =
    [
        "Chat.Read",
        "Chat.ReadBasic",
        "Files.Read.All",
        "User.Read"
    ];

    public static GraphServiceClient CreateClient(string tenantId, string clientId)
    {
        // DeviceCodeCredential prompts the user with a code to enter at aka.ms/devicelogin.
        // This is the recommended interactive flow for CLI tools that cannot open a browser.
        var credential = new DeviceCodeCredential(new DeviceCodeCredentialOptions
        {
            TenantId = tenantId,
            ClientId = clientId,
            DeviceCodeCallback = (info, _) =>
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine();
                Console.WriteLine(info.Message);
                Console.ResetColor();
                return Task.CompletedTask;
            }
        });

        return new GraphServiceClient(credential, Scopes);
    }
}
