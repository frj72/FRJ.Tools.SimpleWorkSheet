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
    public void WithMarkers_None_SetsMarkerStyleToNone()
    {
        var chart = LineChart.Create()
            .WithMarkers(LineChartMarkerStyle.None);

        Assert.Equal(LineChartMarkerStyle.None, chart.MarkerStyle);
    }

    [Fact]
    public void WithMarkers_Diamond_SetsMarkerStyle()
    {
        var chart = LineChart.Create()
            .WithMarkers(LineChartMarkerStyle.Diamond);

        Assert.Equal(LineChartMarkerStyle.Diamond, chart.MarkerStyle);
    }

    [Fact]
    public void WithMarkers_ReturnsSameInstanceWithUpdatedStyle()
    {
        var chart = LineChart.Create();

        var result = chart.WithMarkers(LineChartMarkerStyle.Circle);

        Assert.Same(chart, result);
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
    public void WithTitle_EmptyTitle_ThrowsArgumentException()
    {
        var chart = LineChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(""));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithTitle_WhitespaceTitle_ThrowsArgumentException()
    {
        var chart = LineChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle("   "));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithDataRange_SingleCellCategoriesRange_ThrowsArgumentException()
    {
        var categoriesRange = new CellRange(new(0, 0), new(0, 0));
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = LineChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithDataRange(categoriesRange, valuesRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithDataRange_SingleCellValuesRange_ThrowsArgumentException()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = new CellRange(new(1, 0), new(1, 0));

        var chart = LineChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithDataRange(categoriesRange, valuesRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithPosition_InvalidCoordinates_ThrowsArgumentException()
    {
        var chart = LineChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithPosition(10, 15, 5, 0));
        Assert.Contains("from", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_SetsSize()
    {
        var chart = LineChart.Create()
            .WithSize(8000000, 5000000);

        Assert.Equal(8000000, chart.Size.WidthEmus);
        Assert.Equal(5000000, chart.Size.HeightEmus);
    }

    [Fact]
    public void WithSize_InvalidWidth_ThrowsArgumentException()
    {
        var chart = LineChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(0, 5000000));
        Assert.Contains("width", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_InvalidHeight_ThrowsArgumentException()
    {
        var chart = LineChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(8000000, -1));
        Assert.Contains("height", ex.Message.ToLower());
    }

    [Fact]
    public void AddSeries_AddsSeriesAndReturnsLineChart()
    {
        var dataRange = CellRange.FromBounds(1, 0, 1, 5);
        
        var chart = LineChart.Create()
            .AddSeries("Q1", dataRange);

        Assert.Single(chart.Series);
        Assert.Equal("Q1", chart.Series[0].Name);
        Assert.Equal(dataRange, chart.Series[0].DataRange);
    }

    [Fact]
    public void AddSeries_MultipleSeries_AddsAll()
    {
        var dataRange1 = CellRange.FromBounds(1, 0, 1, 5);
        var dataRange2 = CellRange.FromBounds(2, 0, 2, 5);

        var chart = LineChart.Create()
            .AddSeries("2023", dataRange1)
            .AddSeries("2024", dataRange2);

        Assert.Equal(2, chart.Series.Count);
        Assert.Equal("2023", chart.Series[0].Name);
        Assert.Equal("2024", chart.Series[1].Name);
    }

    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var chart = LineChart.Create();

        Assert.Null(chart.Title);
        Assert.Null(chart.Position);
        Assert.Null(chart.CategoriesRange);
        Assert.Null(chart.ValuesRange);
        Assert.Equal(ChartSize.Default.WidthEmus, chart.Size.WidthEmus);
        Assert.Equal(ChartSize.Default.HeightEmus, chart.Size.HeightEmus);
        Assert.Equal(LineChartMarkerStyle.None, chart.MarkerStyle);
        Assert.False(chart.SmoothLines);
        Assert.Empty(chart.Series);
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
            .WithSize(6000000, 4000000)
            .WithMarkers(LineChartMarkerStyle.Diamond)
            .WithSmoothLines(true);

        Assert.NotNull(chart);
        Assert.Equal("Sales Trend", chart.Title);
        Assert.Equal(categoriesRange, chart.CategoriesRange);
        Assert.Equal(valuesRange, chart.ValuesRange);
        Assert.Equal(6000000, chart.Size.WidthEmus);
        Assert.Equal(LineChartMarkerStyle.Diamond, chart.MarkerStyle);
        Assert.True(chart.SmoothLines);
    }
}
