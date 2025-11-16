using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public class ScatterChart : Chart
{
    public bool ShowTrendline { get; private set; }
    public CellRange? XRange { get; private set; }
    public CellRange? YRange { get; private set; }

    private ScatterChart() : base(ChartType.Scatter)
    {
    }

    public static ScatterChart Create() => new();

    public ScatterChart WithTitle(string? title)
    {
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Title cannot be null or empty", nameof(title));

        Title = title;
        return this;
    }

    public ScatterChart WithXyData(CellRange xRange, CellRange yRange)
    {
        ChartDataRange.ValidateDataRange(xRange);
        ChartDataRange.ValidateDataRange(yRange);

        XRange = xRange;
        YRange = yRange;
        return this;
    }

    public ScatterChart WithPosition(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        Position = new(fromColumn, fromRow, toColumn, toRow);
        return this;
    }

    public ScatterChart WithSize(int width, int height)
    {
        Size = new(width, height);
        return this;
    }

    public ScatterChart WithTrendline(bool show)
    {
        ShowTrendline = show;
        return this;
    }

    public new ScatterChart AddSeries(string name, CellRange dataRange)
    {
        base.AddSeries(name, dataRange);
        return this;
    }
}
