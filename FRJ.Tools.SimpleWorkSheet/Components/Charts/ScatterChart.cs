using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

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

    public ScatterChart WithDataSourceSheet(string sheetName)
    {
        DataSourceSheet = sheetName;
        return this;
    }

    public new ScatterChart AddSeries(string name, CellRange dataRange, string? color = null)
    {
        base.AddSeries(name, dataRange, color);
        return this;
    }

    public ScatterChart WithSeriesColor(string color)
    {
        if (!((string?)color).IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(color));

        SingleSeriesColor = color;
        return this;
    }

    public ScatterChart WithLegendPosition(ChartLegendPosition position)
    {
        LegendPosition = position;
        return this;
    }

    public ScatterChart WithCategoryAxisTitle(string title)
    {
        CategoryAxisTitle = title;
        return this;
    }

    public ScatterChart WithValueAxisTitle(string title)
    {
        ValueAxisTitle = title;
        return this;
    }

    public ScatterChart WithDataLabels(bool show)
    {
        ShowDataLabels = show;
        return this;
    }

    public ScatterChart WithYAxisLabels(bool show)
    {
        ShowYAxisLabels = show;
        return this;
    }

    public ScatterChart WithMajorGridlines(bool show)
    {
        ShowMajorGridlines = show;
        return this;
    }

    public ScatterChart WithSeriesName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Series name cannot be null or empty", nameof(name));

        SingleSeriesName = name;
        return this;
    }
}
