using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class JsonChartBuilder
{
    private readonly JsonWorkbookBuilder _parent;
    private string _chartSheetName = "Chart";
    private string? _categoryColumn;
    private string? _valueColumn;
    private JsonChartType _chartType = JsonChartType.Line;
    private string? _title;
    private string? _categoryAxisTitle;
    private string? _valueAxisTitle;
    private ChartLegendPosition? _legendPosition;

    internal JsonChartBuilder(JsonWorkbookBuilder parent)
    {
        _parent = parent;
    }

    public JsonChartBuilder OnSheet(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Chart sheet name cannot be empty", nameof(sheetName));
        
        _chartSheetName = sheetName;
        return this;
    }

    public JsonChartBuilder UseColumns(string categoryColumn, string valueColumn)
    {
        if (string.IsNullOrWhiteSpace(categoryColumn))
            throw new ArgumentException("Category column cannot be empty", nameof(categoryColumn));
        if (string.IsNullOrWhiteSpace(valueColumn))
            throw new ArgumentException("Value column cannot be empty", nameof(valueColumn));

        _categoryColumn = categoryColumn;
        _valueColumn = valueColumn;
        return this;
    }

    public JsonChartBuilder AsLineChart()
    {
        _chartType = JsonChartType.Line;
        return this;
    }

    public JsonChartBuilder AsBarChart()
    {
        _chartType = JsonChartType.Bar;
        return this;
    }

    public JsonChartBuilder AsAreaChart()
    {
        _chartType = JsonChartType.Area;
        return this;
    }

    public JsonChartBuilder AsPieChart()
    {
        _chartType = JsonChartType.Pie;
        return this;
    }

    public JsonChartBuilder AsScatterChart()
    {
        _chartType = JsonChartType.Scatter;
        return this;
    }

    public JsonChartBuilder WithTitle(string title)
    {
        _title = title;
        return this;
    }

    public JsonChartBuilder WithCategoryAxisTitle(string title)
    {
        _categoryAxisTitle = title;
        return this;
    }

    public JsonChartBuilder WithValueAxisTitle(string title)
    {
        _valueAxisTitle = title;
        return this;
    }

    public JsonChartBuilder WithLegendPosition(ChartLegendPosition position)
    {
        _legendPosition = position;
        return this;
    }

    internal WorkSheet Build(WorkSheet dataSheet)
    {
        if (_categoryColumn == null || _valueColumn == null)
            throw new InvalidOperationException("Must call UseColumns before building chart");

        var categoryRange = _parent.GetColumnRangeByName(_categoryColumn);
        var valueRange = _parent.GetColumnRangeByName(_valueColumn);

        if (categoryRange == null)
            throw new InvalidOperationException($"Column '{_categoryColumn}' not found in data");
        if (valueRange == null)
            throw new InvalidOperationException($"Column '{_valueColumn}' not found in data");

        var chartSheet = new WorkSheet(_chartSheetName);

        Chart chart = _chartType switch
        {
            JsonChartType.Line => CreateLineChart(categoryRange.Value, valueRange.Value),
            JsonChartType.Bar => CreateBarChart(categoryRange.Value, valueRange.Value),
            JsonChartType.Area => CreateAreaChart(categoryRange.Value, valueRange.Value),
            JsonChartType.Pie => CreatePieChart(categoryRange.Value, valueRange.Value),
            JsonChartType.Scatter => CreateScatterChart(categoryRange.Value, valueRange.Value),
            _ => throw new NotSupportedException($"Chart type {_chartType} not supported")
        };

        chartSheet.AddChart(chart);
        return chartSheet;
    }

    private LineChart CreateLineChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = LineChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithPosition(0, 0, 15, 20)
            .WithDataSourceSheet(_parent.DataSheetName);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);

        return chart;
    }

    private BarChart CreateBarChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = BarChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithPosition(0, 0, 15, 20)
            .WithOrientation(BarChartOrientation.Vertical)
            .WithDataSourceSheet(_parent.DataSheetName);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);

        return chart;
    }

    private AreaChart CreateAreaChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = AreaChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithPosition(0, 0, 15, 20)
            .WithDataSourceSheet(_parent.DataSheetName);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);

        return chart;
    }

    private PieChart CreatePieChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = PieChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithPosition(0, 0, 15, 20)
            .WithDataSourceSheet(_parent.DataSheetName);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);

        return chart;
    }

    private ScatterChart CreateScatterChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = ScatterChart.Create()
            .WithXyData(categoryRange, valueRange)
            .WithPosition(0, 0, 15, 20)
            .WithDataSourceSheet(_parent.DataSheetName);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);

        return chart;
    }
}

internal enum JsonChartType
{
    Line,
    Bar,
    Area,
    Pie,
    Scatter
}
