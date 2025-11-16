namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public abstract class Chart
{
    public ChartType Type { get; protected init; }
    public ChartPosition? Position { get; protected set; }
    public ChartSize Size { get; protected set; } = ChartSize.Default;
    public string? Title { get; protected set; }
    public List<ChartSeries> Series { get; } = [];

    protected Chart(ChartType type)
    {
        Type = type;
    }

    public abstract string GetChartTypeName();

    public void AddSeries(string name, Sheet.CellRange dataRange)
    {
        ChartDataRange.ValidateDataRange(dataRange);
        var series = new ChartSeries(name, dataRange);
        Series.Add(series);
    }
}
