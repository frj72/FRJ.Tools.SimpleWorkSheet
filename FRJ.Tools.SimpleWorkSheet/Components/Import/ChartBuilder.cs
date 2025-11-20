using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class ChartBuilder
{
    private readonly WorkbookBuilder _parent;
    private string _chartSheetName = "Chart";
    private string? _categoryColumn;
    private string? _valueColumn;
    private ChartType _chartType = ChartType.Line;
    private string? _title;
    private string? _categoryAxisTitle;
    private string? _valueAxisTitle;
    private ChartLegendPosition? _legendPosition;
    private (uint fromCol, uint fromRow, uint toCol, uint toRow)? _chartPosition;
    private (int width, int height)? _chartSize;
    private bool _showDataLabels;

    internal ChartBuilder(WorkbookBuilder parent) => _parent = parent;

    public ChartBuilder OnSheet(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Chart sheet name cannot be empty", nameof(sheetName));
        
        _chartSheetName = sheetName;
        return this;
    }

    public ChartBuilder UseColumns(string categoryColumn, string valueColumn)
    {
        if (string.IsNullOrWhiteSpace(categoryColumn))
            throw new ArgumentException("Category column cannot be empty", nameof(categoryColumn));
        if (string.IsNullOrWhiteSpace(valueColumn))
            throw new ArgumentException("Value column cannot be empty", nameof(valueColumn));

        _categoryColumn = categoryColumn;
        _valueColumn = valueColumn;
        return this;
    }

    public ChartBuilder AsLineChart()
    {
        _chartType = ChartType.Line;
        return this;
    }

    public ChartBuilder AsBarChart()
    {
        _chartType = ChartType.Bar;
        return this;
    }

    public ChartBuilder AsAreaChart()
    {
        _chartType = ChartType.Area;
        return this;
    }

    public ChartBuilder AsPieChart()
    {
        _chartType = ChartType.Pie;
        return this;
    }

    public ChartBuilder AsScatterChart()
    {
        _chartType = ChartType.Scatter;
        return this;
    }

    public ChartBuilder WithTitle(string title)
    {
        _title = title;
        return this;
    }

    public ChartBuilder WithCategoryAxisTitle(string title)
    {
        _categoryAxisTitle = title;
        return this;
    }

    public ChartBuilder WithValueAxisTitle(string title)
    {
        _valueAxisTitle = title;
        return this;
    }

    public ChartBuilder WithLegendPosition(ChartLegendPosition position)
    {
        _legendPosition = position;
        return this;
    }

    public ChartBuilder WithChartPosition(uint fromCol, uint fromRow, uint toCol, uint toRow)
    {
        _chartPosition = (fromCol, fromRow, toCol, toRow);
        return this;
    }

    public ChartBuilder WithChartSize(int width, int height)
    {
        if (width <= 0)
            throw new ArgumentException("Width must be positive", nameof(width));
        if (height <= 0)
            throw new ArgumentException("Height must be positive", nameof(height));

        _chartSize = (width, height);
        return this;
    }

    public ChartBuilder WithDataLabels(bool show)
    {
        _showDataLabels = show;
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
            ChartType.Line => CreateLineChart(categoryRange.Value, valueRange.Value),
            ChartType.Bar => CreateBarChart(categoryRange.Value, valueRange.Value),
            ChartType.Area => CreateAreaChart(categoryRange.Value, valueRange.Value),
            ChartType.Pie => CreatePieChart(categoryRange.Value, valueRange.Value),
            ChartType.Scatter => CreateScatterChart(categoryRange.Value, valueRange.Value),
            _ => throw new NotSupportedException($"Chart type {_chartType} not supported")
        };

        chartSheet.AddChart(chart);
        return chartSheet;
    }

    private LineChart CreateLineChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = LineChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithDataSourceSheet(_parent.DataSheetName);

        chart = _chartPosition != null ? chart.WithPosition(_chartPosition.Value.fromCol, _chartPosition.Value.fromRow, _chartPosition.Value.toCol, _chartPosition.Value.toRow) : chart.WithPosition(0, 0, 15, 20);

        if (_chartSize != null)
            chart = chart.WithSize(_chartSize.Value.width, _chartSize.Value.height);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);
        if (_showDataLabels)
            chart = chart.WithDataLabels(true);

        return chart;
    }

    private BarChart CreateBarChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = BarChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithOrientation(BarChartOrientation.Vertical)
            .WithDataSourceSheet(_parent.DataSheetName);

        chart = _chartPosition != null ? chart.WithPosition(_chartPosition.Value.fromCol, _chartPosition.Value.fromRow, _chartPosition.Value.toCol, _chartPosition.Value.toRow) : chart.WithPosition(0, 0, 15, 20);

        if (_chartSize != null)
            chart = chart.WithSize(_chartSize.Value.width, _chartSize.Value.height);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);
        if (_showDataLabels)
            chart = chart.WithDataLabels(true);

        return chart;
    }

    private AreaChart CreateAreaChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = AreaChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithDataSourceSheet(_parent.DataSheetName);

        chart = _chartPosition != null ? chart.WithPosition(_chartPosition.Value.fromCol, _chartPosition.Value.fromRow, _chartPosition.Value.toCol, _chartPosition.Value.toRow) : chart.WithPosition(0, 0, 15, 20);

        if (_chartSize != null)
            chart = chart.WithSize(_chartSize.Value.width, _chartSize.Value.height);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);
        if (_showDataLabels)
            chart = chart.WithDataLabels(true);

        return chart;
    }

    private PieChart CreatePieChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = PieChart.Create()
            .WithDataRange(categoryRange, valueRange)
            .WithDataSourceSheet(_parent.DataSheetName);

        chart = _chartPosition != null ? chart.WithPosition(_chartPosition.Value.fromCol, _chartPosition.Value.fromRow, _chartPosition.Value.toCol, _chartPosition.Value.toRow) : chart.WithPosition(0, 0, 15, 20);

        if (_chartSize != null)
            chart = chart.WithSize(_chartSize.Value.width, _chartSize.Value.height);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);
        if (_showDataLabels)
            chart = chart.WithDataLabels(true);

        return chart;
    }

    private ScatterChart CreateScatterChart(CellRange categoryRange, CellRange valueRange)
    {
        var chart = ScatterChart.Create()
            .WithXyData(categoryRange, valueRange)
            .WithDataSourceSheet(_parent.DataSheetName);

        chart = _chartPosition != null ? chart.WithPosition(_chartPosition.Value.fromCol, _chartPosition.Value.fromRow, _chartPosition.Value.toCol, _chartPosition.Value.toRow) : chart.WithPosition(0, 0, 15, 20);

        if (_chartSize != null)
            chart = chart.WithSize(_chartSize.Value.width, _chartSize.Value.height);

        if (_title != null)
            chart = chart.WithTitle(_title);
        if (_categoryAxisTitle != null)
            chart = chart.WithCategoryAxisTitle(_categoryAxisTitle);
        if (_valueAxisTitle != null)
            chart = chart.WithValueAxisTitle(_valueAxisTitle);
        if (_legendPosition != null)
            chart = chart.WithLegendPosition(_legendPosition.Value);
        if (_showDataLabels)
            chart = chart.WithDataLabels(true);

        return chart;
    }
}