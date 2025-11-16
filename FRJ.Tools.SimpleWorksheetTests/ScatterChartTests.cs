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
    public void WithTitle_EmptyTitle_ThrowsArgumentException()
    {
        var chart = ScatterChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(""));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithTitle_WhitespaceTitle_ThrowsArgumentException()
    {
        var chart = ScatterChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle("   "));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithXyData_SingleCellXRange_ThrowsArgumentException()
    {
        var xRange = new CellRange(new(0, 0), new(0, 0));
        var yRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = ScatterChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithXyData(xRange, yRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithXyData_SingleCellYRange_ThrowsArgumentException()
    {
        var xRange = CellRange.FromBounds(0, 0, 0, 5);
        var yRange = new CellRange(new(1, 0), new(1, 0));

        var chart = ScatterChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithXyData(xRange, yRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithPosition_InvalidCoordinates_ThrowsArgumentException()
    {
        var chart = ScatterChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithPosition(10, 15, 5, 0));
        Assert.Contains("from", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_SetsSize()
    {
        var chart = ScatterChart.Create()
            .WithSize(8000000, 5000000);

        Assert.Equal(8000000, chart.Size.WidthEmus);
        Assert.Equal(5000000, chart.Size.HeightEmus);
    }

    [Fact]
    public void WithSize_InvalidWidth_ThrowsArgumentException()
    {
        var chart = ScatterChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(0, 5000000));
        Assert.Contains("width", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_InvalidHeight_ThrowsArgumentException()
    {
        var chart = ScatterChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(8000000, -1));
        Assert.Contains("height", ex.Message.ToLower());
    }

    [Fact]
    public void AddSeries_AddsSeriesAndReturnsScatterChart()
    {
        var dataRange = CellRange.FromBounds(1, 0, 1, 5);
        
        var chart = ScatterChart.Create()
            .AddSeries("Data Set 1", dataRange);

        Assert.Single(chart.Series);
        Assert.Equal("Data Set 1", chart.Series[0].Name);
        Assert.Equal(dataRange, chart.Series[0].DataRange);
    }

    [Fact]
    public void AddSeries_MultipleSeries_AddsAll()
    {
        var dataRange1 = CellRange.FromBounds(1, 0, 1, 5);
        var dataRange2 = CellRange.FromBounds(2, 0, 2, 5);

        var chart = ScatterChart.Create()
            .AddSeries("Set A", dataRange1)
            .AddSeries("Set B", dataRange2);

        Assert.Equal(2, chart.Series.Count);
        Assert.Equal("Set A", chart.Series[0].Name);
        Assert.Equal("Set B", chart.Series[1].Name);
    }

    [Fact]
    public void WithDataSourceSheet_SetsDataSourceSheet()
    {
        var chart = ScatterChart.Create()
            .WithDataSourceSheet("SourceSheet");

        Assert.Equal("SourceSheet", chart.DataSourceSheet);
    }

    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var chart = ScatterChart.Create();

        Assert.Null(chart.Title);
        Assert.Null(chart.Position);
        Assert.Null(chart.XRange);
        Assert.Null(chart.YRange);
        Assert.Null(chart.DataSourceSheet);
        Assert.Equal(ChartSize.Default.WidthEmus, chart.Size.WidthEmus);
        Assert.Equal(ChartSize.Default.HeightEmus, chart.Size.HeightEmus);
        Assert.False(chart.ShowTrendline);
        Assert.Empty(chart.Series);
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
            .WithSize(6000000, 4000000)
            .WithTrendline(true);

        Assert.NotNull(chart);
        Assert.Equal("X vs Y", chart.Title);
        Assert.Equal(xRange, chart.XRange);
        Assert.Equal(yRange, chart.YRange);
        Assert.Equal(6000000, chart.Size.WidthEmus);
        Assert.True(chart.ShowTrendline);
    }
}
