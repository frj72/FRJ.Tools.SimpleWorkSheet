using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartEdgeCasesTests
{
    [Fact]
    public void BarChart_WithVeryLongTitle_HandlesCorrectly()
    {
        var longTitle = new string('A', 255);
        
        var chart = BarChart.Create()
            .WithTitle(longTitle);
        
        Assert.Equal(longTitle, chart.Title);
    }

    [Fact]
    public void LineChart_WithUnicodeTitle_PreservesUnicode()
    {
        const string unicodeTitle = "Chart 图表 📊 مخطط";
        
        var chart = LineChart.Create()
            .WithTitle(unicodeTitle);
        
        Assert.Equal(unicodeTitle, chart.Title);
    }

    [Fact]
    public void PieChart_WithMaxExplosion_SetsCorrectly()
    {
        var chart = PieChart.Create()
            .WithExplosion(100);
        
        Assert.Equal(100u, chart.Explosion);
    }

    [Fact]
    public void PieChart_WithMinExplosion_SetsCorrectly()
    {
        var chart = PieChart.Create()
            .WithExplosion(0);
        
        Assert.Equal(0u, chart.Explosion);
    }

    [Fact]
    public void PieChart_WithMaxFirstSliceAngle_SetsCorrectly()
    {
        var chart = PieChart.Create()
            .WithFirstSliceAngle(360);
        
        Assert.Equal(360u, chart.FirstSliceAngle);
    }

    [Fact]
    public void PieChart_WithMinFirstSliceAngle_SetsCorrectly()
    {
        var chart = PieChart.Create()
            .WithFirstSliceAngle(0);
        
        Assert.Equal(0u, chart.FirstSliceAngle);
    }

    [Fact]
    public void ScatterChart_WithVerySmallDataRange_HandlesCorrectly()
    {
        var smallRange = CellRange.FromBounds(0, 0, 0, 1);
        
        var chart = ScatterChart.Create()
            .WithXyData(smallRange, smallRange);
        
        Assert.NotNull(chart);
    }

    [Fact]
    public void ScatterChart_WithVeryLargeDataRange_HandlesCorrectly()
    {
        var largeRange = CellRange.FromBounds(0, 0, 0, 10000);
        
        var chart = ScatterChart.Create()
            .WithXyData(largeRange, largeRange);
        
        Assert.NotNull(chart);
    }

    [Fact]
    public void AreaChart_WithDefaultSettings_CreatesChart()
    {
        var chart = AreaChart.Create();
        
        Assert.NotNull(chart);
        Assert.Null(chart.Title);
    }

    [Fact]
    public void AreaChart_WithAllProperties_SetsCorrectly()
    {
        const string title = "Area Chart";
        var dataRange = CellRange.FromBounds(0, 0, 0, 10);
        
        var chart = AreaChart.Create()
            .WithTitle(title)
            .WithDataRange(dataRange, dataRange)
            .WithSize(800, 600)
            .WithPosition(0, 0, 10, 15);
        
        Assert.Equal(title, chart.Title);
    }

    [Fact]
    public void BarChart_WithHorizontalOrientation_SetsCorrectly()
    {
        var chart = BarChart.Create()
            .WithOrientation(BarChartOrientation.Horizontal);
        
        Assert.Equal(BarChartOrientation.Horizontal, chart.Orientation);
    }

    [Fact]
    public void LineChart_WithSmoothLines_SetsCorrectly()
    {
        var chart = LineChart.Create()
            .WithSmoothLines(true);
        
        Assert.True(chart.SmoothLines);
    }

    [Fact]
    public void LineChart_WithDiamondMarkers_SetsCorrectly()
    {
        var chart = LineChart.Create()
            .WithMarkers(LineChartMarkerStyle.Diamond);
        
        Assert.Equal(LineChartMarkerStyle.Diamond, chart.MarkerStyle);
    }

    [Fact]
    public void ScatterChart_WithTrendline_SetsCorrectly()
    {
        var chart = ScatterChart.Create()
            .WithTrendline(true);
        
        Assert.True(chart.ShowTrendline);
    }

    [Fact]
    public void Chart_WithMaxSizeValues_HandlesCorrectly()
    {
        var chart = BarChart.Create()
            .WithSize(int.MaxValue, int.MaxValue);
        
        Assert.NotNull(chart);
    }

    [Fact]
    public void Chart_WithLargePosition_HandlesCorrectly()
    {
        var chart = BarChart.Create()
            .WithPosition(1000, 1000, 2000, 2000);
        
        Assert.NotNull(chart);
    }

    [Fact]
    public void Chart_WithMultipleSeries_AllAdded()
    {
        var range1 = CellRange.FromBounds(0, 0, 0, 5);
        var range2 = CellRange.FromBounds(0, 6, 0, 11);
        
        var chart = BarChart.Create()
            .AddSeries("Series1", range1)
            .AddSeries("Series2", range2);
        
        Assert.Equal(2, chart.Series.Count);
    }

    [Fact]
    public void Chart_WithEmptySeriesName_ThrowsException()
    {
        var range = CellRange.FromBounds(0, 0, 0, 5);
        
        Assert.Throws<ArgumentException>(() => 
            BarChart.Create().AddSeries("", range));
    }
}
