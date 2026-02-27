using TeamsChatCrawler;

var options = CliOptions.Parse(args);
if (options.ShowHelp)
{
    CliOptions.PrintHelp();
    return;
}

var token = Environment.GetEnvironmentVariable("TEAMS_ACCESS_TOKEN");
if (string.IsNullOrWhiteSpace(token))
{
    Console.Error.WriteLine("Missing TEAMS_ACCESS_TOKEN environment variable.");
    Console.Error.WriteLine("Provide a Microsoft Graph delegated access token with Chat.Read and Files.Read permissions.");
    return;
}

using var httpClient = new HttpClient
{
    BaseAddress = new Uri("https://graph.microsoft.com/v1.0/")
};
httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

var outputRoot = Path.GetFullPath(options.OutputDirectory);
Directory.CreateDirectory(outputRoot);

var graphClient = new GraphTeamsClient(httpClient);
var crawler = new TeamsCrawler(graphClient);

try
{
    var summary = await crawler.RunAsync(options, outputRoot, CancellationToken.None);

    Console.WriteLine("Crawl completed.");
    Console.WriteLine($"Chats scanned: {summary.ChatsScanned}");
    Console.WriteLine($"Messages exported: {summary.MessagesExported}");
    Console.WriteLine($"Attachments downloaded: {summary.AttachmentsDownloaded}");
    Console.WriteLine($"Output: {outputRoot}");
}
catch (Exception ex)
{
    Console.Error.WriteLine("Crawler failed:");
    Console.Error.WriteLine(ex.Message);
}
