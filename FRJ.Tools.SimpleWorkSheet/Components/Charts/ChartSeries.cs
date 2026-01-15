using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public record ChartSeries
{
    public string Name { get; init; }
    public CellRange DataRange { get; init; }
    public string? Color { get; init; }

    public ChartSeries(string? name, CellRange dataRange, string? color = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Series name cannot be null or whitespace", nameof(name));

        if (color != null && !color.IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(color));

        Name = name;
        DataRange = dataRange;
        Color = color;
    }
}
