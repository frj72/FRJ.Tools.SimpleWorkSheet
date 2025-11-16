using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ScatterChartTests
{
    [Fact]
    public void Create_CreatesScatterChartInstance()
    {
        var chart = ScatterChart.Create();

        Assert.NotNull(chart);
        Assert.Equal(ChartType.Scatter, chart.Type);
    }

    [Fact]
    public void GetChartTypeName_ReturnsScatterChart()
    {
        var chart = ScatterChart.Create();

        Assert.Equal("scatterChart", chart.GetChartTypeName());
    }

    [Fact]
    public void DefaultShowTrendline_IsFalse()
    {
        var chart = ScatterChart.Create();

        Assert.False(chart.ShowTrendline);
    }

    [Fact]
    public void WithTitle_SetsTitle()
    {
        var chart = ScatterChart.Create()
            .WithTitle("Correlation Analysis");

        Assert.Equal("Correlation Analysis", chart.Title);
    }

    [Fact]
    public void WithTitle_NullTitle_ThrowsArgumentException()
    {
        var chart = ScatterChart.Create();
        string? nullTitle = null;

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(nullTitle));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithXyData_SetsRanges()
    {
        var xRange = CellRange.FromBounds(0, 0, 0, 5);
        var yRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = ScatterChart.Create()
            .WithXyData(xRange, yRange);

        Assert.Equal(xRange, chart.XRange);
        Assert.Equal(yRange, chart.YRange);
    }

    [Fact]
    public void WithTrendline_SetsTrendline()
    {
        var chart = ScatterChart.Create()
            .WithTrendline(true);

        Assert.True(chart.ShowTrendline);
    }

    [Fact]
    public void WithPosition_SetsPosition()
    {
        var chart = ScatterChart.Create()
            .WithPosition(5, 0, 10, 15);

        Assert.NotNull(chart.Position);
        Assert.Equal(5u, chart.Position.FromColumn);
        Assert.Equal(0u, chart.Position.FromRow);
        Assert.Equal(10u, chart.Position.ToColumn);
        Assert.Equal(15u, chart.Position.ToRow);
    }

    [Fact]
    public void FluentAPI_AllMethods_ReturnScatterChart()
    {
        var xRange = CellRange.FromBounds(0, 0, 0, 5);
        var yRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = ScatterChart.Create()
            .WithTitle("X vs Y")
            .WithXyData(xRange, yRange)
            .WithPosition(5, 0, 10, 15)
            .WithTrendline(true);

        Assert.NotNull(chart);
        Assert.Equal("X vs Y", chart.Title);
        Assert.Equal(xRange, chart.XRange);
        Assert.Equal(yRange, chart.YRange);
        Assert.True(chart.ShowTrendline);
    }
}
