using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartTests
{
    private class TestChart : Chart
    {
        public TestChart() : base(ChartType.Bar) { }
    }

    [Fact]
    public void Chart_Constructor_SetsType()
    {
        var chart = new TestChart();

        Assert.Equal(ChartType.Bar, chart.Type);
    }

    [Fact]
    public void Chart_DefaultSize_IsSet()
    {
        var chart = new TestChart();

        Assert.NotNull(chart.Size);
        Assert.Equal(ChartSize.Default.WidthEmus, chart.Size.WidthEmus);
        Assert.Equal(ChartSize.Default.HeightEmus, chart.Size.HeightEmus);
    }

    [Fact]
    public void Chart_Series_StartsEmpty()
    {
        var chart = new TestChart();

        Assert.Empty(chart.Series);
    }

    [Fact]
    public void Chart_AddSeries_AddsSeriesToCollection()
    {
        var chart = new TestChart();
        var range = CellRange.FromBounds(0, 0, 0, 5);

        chart.AddSeries("Test Series", range);

        Assert.Single(chart.Series);
        Assert.Equal("Test Series", chart.Series[0].Name);
        Assert.Equal(range, chart.Series[0].DataRange);
    }

    [Fact]
    public void Chart_AddMultipleSeries_AddsAll()
    {
        var chart = new TestChart();
        var range1 = CellRange.FromBounds(0, 0, 0, 5);
        var range2 = CellRange.FromBounds(1, 0, 1, 5);

        chart.AddSeries("Series 1", range1);
        chart.AddSeries("Series 2", range2);

        Assert.Equal(2, chart.Series.Count);
        Assert.Equal("Series 1", chart.Series[0].Name);
        Assert.Equal("Series 2", chart.Series[1].Name);
    }

    [Fact]
    public void Chart_AddSeries_ThrowsOnSingleCellRange()
    {
        var chart = new TestChart();
        var singleCell = CellRange.FromBounds(0, 0, 0, 0);

        Assert.Throws<ArgumentException>(() => chart.AddSeries("Test", singleCell));
    }
}
