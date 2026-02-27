using System.CommandLine;
using Microsoft.Graph;
using TeamsChatCrawler.Models;
using TeamsChatCrawler.Services;

// ---------------------------------------------------------------------------
// Root command definition
// ---------------------------------------------------------------------------
var rootCommand = new RootCommand(
    "Teams Chat Crawler – exports messages, pictures and documents from Microsoft Teams chats.");

var dateOption = new Option<string?>(
    name: "--date",
    description: "Reference date (yyyy-MM-dd). Defaults to today. " +
                 "The tool fetches the 5 days prior to this date.")
{
    IsRequired = false
};
dateOption.AddAlias("-d");

var chatsOption = new Option<string[]>(
    name: "--chats",
    description: "One or more chat names, topics or member names to filter. " +
                 "Separate multiple values with spaces after the flag. " +
                 "If omitted, ALL chats are crawled.")
{
    IsRequired = false,
    AllowMultipleArgumentsPerToken = true
};
chatsOption.AddAlias("-c");

var outputOption = new Option<string>(
    name: "--output",
    description: "Root directory where downloaded content will be saved.",
    getDefaultValue: () => Path.Combine(Directory.GetCurrentDirectory(), "teams-export"))
{
    IsRequired = false
};
outputOption.AddAlias("-o");

var tenantOption = new Option<string?>(
    name: "--tenant-id",
    description: "Azure AD tenant ID (or 'common' for multi-tenant apps). " +
                 "Can also be set via TEAMS_TENANT_ID environment variable.")
{
    IsRequired = false
};
tenantOption.AddAlias("-t");

var clientOption = new Option<string?>(
    name: "--client-id",
    description: "Azure AD app registration client ID. " +
                 "Can also be set via TEAMS_CLIENT_ID environment variable.")
{
    IsRequired = false
};
clientOption.AddAlias("-i");

rootCommand.AddOption(dateOption);
rootCommand.AddOption(chatsOption);
rootCommand.AddOption(outputOption);
rootCommand.AddOption(tenantOption);
rootCommand.AddOption(clientOption);

// ---------------------------------------------------------------------------
// Handler
// ---------------------------------------------------------------------------
rootCommand.SetHandler(async (
    string? dateStr,
    string[] chats,
    string output,
    string? tenantIdArg,
    string? clientIdArg) =>
{
    // -- Resolve tenant / client IDs (arg takes precedence over env var)
    var tenantId = tenantIdArg ?? Environment.GetEnvironmentVariable("TEAMS_TENANT_ID");
    var clientId = clientIdArg ?? Environment.GetEnvironmentVariable("TEAMS_CLIENT_ID");

    if (string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(clientId))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Error.WriteLine(
            "ERROR: --tenant-id and --client-id are required (or set TEAMS_TENANT_ID / TEAMS_CLIENT_ID).");
        Console.ResetColor();
        Environment.Exit(1);
    }

    // -- Parse reference date
    DateOnly referenceDate;
    if (string.IsNullOrWhiteSpace(dateStr))
    {
        referenceDate = DateOnly.FromDateTime(DateTime.Today);
    }
    else if (!DateOnly.TryParseExact(dateStr, "yyyy-MM-dd",
                 System.Globalization.CultureInfo.InvariantCulture,
                 System.Globalization.DateTimeStyles.None,
                 out referenceDate))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Error.WriteLine($"ERROR: Invalid date '{dateStr}'. Expected format: yyyy-MM-dd.");
        Console.ResetColor();
        Environment.Exit(1);
    }

    var options = new CrawlOptions
    {
        ReferenceDate = referenceDate,
        DaysLookback  = 5,
        ChatFilters   = chats.ToList()
    };

    // -- Build Graph client
    GraphServiceClient graph;
    try
    {
        graph = AuthService.CreateClient(tenantId!, clientId!);
    }
    catch (Exception ex)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Error.WriteLine($"ERROR: Failed to create Graph client: {ex.Message}");
        Console.ResetColor();
        Environment.Exit(1);
        return;
    }

    // -- Compose services
    using var httpClient = new HttpClient();
    var chatService      = new ChatService(graph);
    var messageService   = new MessageService(graph);
    var downloadService  = new DownloadService(graph, httpClient);
    var orchestrator     = new CrawlerOrchestrator(chatService, messageService, downloadService);

    var cts = new CancellationTokenSource();
    Console.CancelKeyPress += (_, e) =>
    {
        e.Cancel = true;
        Console.WriteLine("\nCancelling...");
        cts.Cancel();
    };

    try
    {
        await orchestrator.RunAsync(options, output, cts.Token);
    }
    catch (OperationCanceledException)
    {
        Console.WriteLine("Operation cancelled by user.");
    }
    catch (Exception ex)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Error.WriteLine($"\nERROR: {ex.Message}");
        Console.ResetColor();
        Environment.Exit(1);
    }
},
dateOption, chatsOption, outputOption, tenantOption, clientOption);

return await rootCommand.InvokeAsync(args);
