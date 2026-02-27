using System.Globalization;

namespace TeamsChatCrawler;

public sealed record CliOptions(
    DateOnly AnchorDate,
    IReadOnlyList<string> ChatFilters,
    IReadOnlyList<string> GroupFilters,
    string? GroupConfigPath,
    string OutputDirectory,
    bool ShowHelp)
{
    public static CliOptions Parse(string[] args)
    {
        var anchorDate = DateOnly.FromDateTime(DateTime.UtcNow);
        var chatFilters = new List<string>();
        var groupFilters = new List<string>();
        string? groupConfigPath = null;
        var outputDirectory = "output";
        var showHelp = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg)
            {
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                case "-d":
                case "--date":
                    anchorDate = ParseDate(RequireValue(args, ref i, arg));
                    break;
                case "-c":
                case "--chat":
                    chatFilters.AddRange(ParseCsv(RequireValue(args, ref i, arg)));
                    break;
                case "-g":
                case "--group":
                    groupFilters.AddRange(ParseCsv(RequireValue(args, ref i, arg)));
                    break;
                case "--group-config":
                    groupConfigPath = RequireValue(args, ref i, arg);
                    break;
                case "-o":
                case "--output":
                    outputDirectory = RequireValue(args, ref i, arg);
                    break;
                default:
                    throw new ArgumentException($"Unknown argument '{arg}'. Use --help for usage.");
            }
        }

        return new CliOptions(anchorDate, chatFilters, groupFilters, groupConfigPath, outputDirectory, showHelp);
    }

    public static void PrintHelp()
    {
        Console.WriteLine("TeamsChatCrawler - Export Teams chat messages and files from a 5-day window");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  TeamsChatCrawler [--date yyyy-MM-dd] [--chat names] [--group names] [--group-config path] [--output path]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -d, --date           Anchor date in yyyy-MM-dd (default: today UTC)");
        Console.WriteLine("  -c, --chat           Comma-separated chat topic/title filters");
        Console.WriteLine("  -g, --group          Comma-separated chat grouping names");
        Console.WriteLine("      --group-config   JSON file that maps group names to chatIds/chatTitles");
        Console.WriteLine("  -o, --output         Output directory (default: ./output)");
        Console.WriteLine("  -h, --help           Show help");
    }

    private static string RequireValue(string[] args, ref int i, string arg)
    {
        if (i + 1 >= args.Length)
        {
            throw new ArgumentException($"Missing value after {arg}");
        }

        i++;
        return args[i];
    }

    private static DateOnly ParseDate(string value)
    {
        if (!DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
        {
            throw new ArgumentException($"Invalid date '{value}'. Expected yyyy-MM-dd.");
        }

        return parsed;
    }

    private static IEnumerable<string> ParseCsv(string value) =>
        value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}
