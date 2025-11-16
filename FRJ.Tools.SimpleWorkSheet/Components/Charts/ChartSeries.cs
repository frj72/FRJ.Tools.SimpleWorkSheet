using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public record ChartSeries
{
    public string Name { get; init; }
    public CellRange DataRange { get; init; }

    public ChartSeries(string? name, CellRange dataRange)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Series name cannot be null or whitespace", nameof(name));

        Name = name;
        DataRange = dataRange;
    }
}
