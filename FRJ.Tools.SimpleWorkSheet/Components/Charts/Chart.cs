using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public abstract class Chart
{
    public ChartType Type { get; protected init; }
    public ChartPosition? Position { get; protected set; }
    public ChartSize Size { get; protected set; } = ChartSize.Default;
    public string? Title { get; protected set; }
    public string? DataSourceSheet { get; protected set; }
    public List<ChartSeries> Series { get; } = [];
    public ChartLegendPosition LegendPosition { get; protected set; } = ChartLegendPosition.Right;
    public string? CategoryAxisTitle { get; protected set; }
    public string? ValueAxisTitle { get; protected set; }
    public bool ShowDataLabels { get; protected set; }
    public bool ShowYAxisLabels { get; protected set; } = true;
    public bool ShowMajorGridlines { get; protected set; } = true;
    internal string? SingleSeriesName { get; set; }
    internal string? SingleSeriesColor { get; set; }

    protected Chart(ChartType type) => Type = type;

    public void AddSeries(string name, CellRange dataRange, string? color = null)
    {
        ChartDataRange.ValidateDataRange(dataRange);
        var series = new ChartSeries(name, dataRange, color);
        Series.Add(series);
    }
}
