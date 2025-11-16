using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class LineChartTests
{
    [Fact]
    public void Create_CreatesLineChartInstance()
    {
        var chart = LineChart.Create();

        Assert.NotNull(chart);
        Assert.Equal(ChartType.Line, chart.Type);
    }

    [Fact]
    public void GetChartTypeName_ReturnsLineChart()
    {
        var chart = LineChart.Create();

        Assert.Equal("lineChart", chart.GetChartTypeName());
    }

    [Fact]
    public void DefaultMarkerStyle_IsNone()
    {
        var chart = LineChart.Create();

        Assert.Equal(LineChartMarkerStyle.None, chart.MarkerStyle);
    }

    [Fact]
    public void DefaultSmoothLines_IsFalse()
    {
        var chart = LineChart.Create();

        Assert.False(chart.SmoothLines);
    }

    [Fact]
    public void WithTitle_SetsTitle()
    {
        var chart = LineChart.Create()
            .WithTitle("Trend Chart");

        Assert.Equal("Trend Chart", chart.Title);
    }

    [Fact]
    public void WithTitle_NullTitle_ThrowsArgumentException()
    {
        var chart = LineChart.Create();
        string? nullTitle = null;

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(nullTitle));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithDataRange_SetsRanges()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = LineChart.Create()
            .WithDataRange(categoriesRange, valuesRange);

        Assert.Equal(categoriesRange, chart.CategoriesRange);
        Assert.Equal(valuesRange, chart.ValuesRange);
    }

    [Fact]
    public void WithMarkers_SetsMarkerStyle()
    {
        var chart = LineChart.Create()
            .WithMarkers(LineChartMarkerStyle.Circle);

        Assert.Equal(LineChartMarkerStyle.Circle, chart.MarkerStyle);
    }

    [Fact]
    public void WithSmoothLines_SetsSmoothLines()
    {
        var chart = LineChart.Create()
            .WithSmoothLines(true);

        Assert.True(chart.SmoothLines);
    }

    [Fact]
    public void WithPosition_SetsPosition()
    {
        var chart = LineChart.Create()
            .WithPosition(5, 0, 10, 15);

        Assert.NotNull(chart.Position);
        Assert.Equal(5u, chart.Position.FromColumn);
        Assert.Equal(0u, chart.Position.FromRow);
        Assert.Equal(10u, chart.Position.ToColumn);
        Assert.Equal(15u, chart.Position.ToRow);
    }

    [Fact]
    public void FluentAPI_AllMethods_ReturnLineChart()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = LineChart.Create()
            .WithTitle("Sales Trend")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(5, 0, 10, 15)
            .WithMarkers(LineChartMarkerStyle.Diamond)
            .WithSmoothLines(true);

        Assert.NotNull(chart);
        Assert.Equal("Sales Trend", chart.Title);
        Assert.Equal(categoriesRange, chart.CategoriesRange);
        Assert.Equal(valuesRange, chart.ValuesRange);
        Assert.Equal(LineChartMarkerStyle.Diamond, chart.MarkerStyle);
        Assert.True(chart.SmoothLines);
    }
}
