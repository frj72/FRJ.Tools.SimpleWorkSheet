using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public class LineChart : Chart
{
    public LineChartMarkerStyle MarkerStyle { get; private set; } = LineChartMarkerStyle.None;
    public bool SmoothLines { get; private set; }
    public CellRange? CategoriesRange { get; private set; }
    public CellRange? ValuesRange { get; private set; }

    private LineChart() : base(ChartType.Line)
    {
    }

    public static LineChart Create() => new();

    public override string GetChartTypeName() => "lineChart";

    public LineChart WithTitle(string? title)
    {
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Title cannot be null or empty", nameof(title));

        Title = title;
        return this;
    }

    public LineChart WithDataRange(CellRange categoriesRange, CellRange valuesRange)
    {
        ChartDataRange.ValidateDataRange(categoriesRange);
        ChartDataRange.ValidateDataRange(valuesRange);

        CategoriesRange = categoriesRange;
        ValuesRange = valuesRange;
        return this;
    }

    public LineChart WithPosition(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        Position = new(fromColumn, fromRow, toColumn, toRow);
        return this;
    }

    public LineChart WithSize(int width, int height)
    {
        Size = new(width, height);
        return this;
    }

    public LineChart WithMarkers(LineChartMarkerStyle style)
    {
        MarkerStyle = style;
        return this;
    }

    public LineChart WithSmoothLines(bool smooth)
    {
        SmoothLines = smooth;
        return this;
    }

    public new LineChart AddSeries(string name, CellRange dataRange)
    {
        base.AddSeries(name, dataRange);
        return this;
    }
}
