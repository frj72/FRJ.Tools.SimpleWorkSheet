using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class PieChartTests
{
    [Fact]
    public void Create_CreatesPieChartInstance()
    {
        var chart = PieChart.Create();

        Assert.NotNull(chart);
        Assert.Equal(ChartType.Pie, chart.Type);
    }

    [Fact]
    public void GetChartTypeName_ReturnsPieChart()
    {
        var chart = PieChart.Create();

        Assert.Equal("pieChart", chart.GetChartTypeName());
    }

    [Fact]
    public void DefaultExplosion_IsZero()
    {
        var chart = PieChart.Create();

        Assert.Equal(0u, chart.Explosion);
    }

    [Fact]
    public void DefaultFirstSliceAngle_IsZero()
    {
        var chart = PieChart.Create();

        Assert.Equal(0u, chart.FirstSliceAngle);
    }

    [Fact]
    public void WithTitle_SetsTitle()
    {
        var chart = PieChart.Create()
            .WithTitle("Market Share");

        Assert.Equal("Market Share", chart.Title);
    }

    [Fact]
    public void WithTitle_NullTitle_ThrowsArgumentException()
    {
        var chart = PieChart.Create();
        string? nullTitle = null;

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(nullTitle));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithDataRange_SetsRanges()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = PieChart.Create()
            .WithDataRange(categoriesRange, valuesRange);

        Assert.Equal(categoriesRange, chart.CategoriesRange);
        Assert.Equal(valuesRange, chart.ValuesRange);
    }

    [Fact]
    public void WithExplosion_SetsExplosion()
    {
        var chart = PieChart.Create()
            .WithExplosion(25);

        Assert.Equal(25u, chart.Explosion);
    }

    [Fact]
    public void WithExplosion_Over100_ThrowsArgumentException()
    {
        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithExplosion(101));
        Assert.Contains("percentage", ex.Message.ToLower());
    }

    [Fact]
    public void WithFirstSliceAngle_SetsAngle()
    {
        var chart = PieChart.Create()
            .WithFirstSliceAngle(90);

        Assert.Equal(90u, chart.FirstSliceAngle);
    }

    [Fact]
    public void WithFirstSliceAngle_Over360_ThrowsArgumentException()
    {
        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithFirstSliceAngle(361));
        Assert.Contains("degrees", ex.Message.ToLower());
    }

    [Fact]
    public void WithPosition_SetsPosition()
    {
        var chart = PieChart.Create()
            .WithPosition(5, 0, 10, 15);

        Assert.NotNull(chart.Position);
        Assert.Equal(5u, chart.Position.FromColumn);
    }

    [Fact]
    public void FluentAPI_AllMethods_ReturnPieChart()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = PieChart.Create()
            .WithTitle("Market Share")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(5, 0, 10, 15)
            .WithExplosion(10)
            .WithFirstSliceAngle(45);

        Assert.NotNull(chart);
        Assert.Equal("Market Share", chart.Title);
        Assert.Equal(10u, chart.Explosion);
        Assert.Equal(45u, chart.FirstSliceAngle);
    }
}
