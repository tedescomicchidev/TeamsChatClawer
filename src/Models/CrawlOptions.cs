namespace TeamsChatCrawler.Models;

public class CrawlOptions
{
    public DateOnly ReferenceDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public int DaysLookback { get; set; } = 5;
    public List<string> ChatFilters { get; set; } = new();

    public DateTimeOffset WindowStart =>
        new DateTimeOffset(ReferenceDate.Year, ReferenceDate.Month, ReferenceDate.Day, 0, 0, 0, TimeSpan.Zero)
            .AddDays(-DaysLookback);

    public DateTimeOffset WindowEnd =>
        new DateTimeOffset(ReferenceDate.Year, ReferenceDate.Month, ReferenceDate.Day, 23, 59, 59, TimeSpan.Zero);
}
