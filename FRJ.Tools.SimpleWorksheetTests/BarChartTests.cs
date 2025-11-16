using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class BarChartTests
{
    [Fact]
    public void Create_CreatesBarChartInstance()
    {
        var chart = BarChart.Create();

        Assert.NotNull(chart);
        Assert.Equal(ChartType.Bar, chart.Type);
    }

    [Fact]
    public void DefaultOrientation_IsVertical()
    {
        var chart = BarChart.Create();

        Assert.Equal(BarChartOrientation.Vertical, chart.Orientation);
    }

    [Fact]
    public void WithTitle_SetsTitle()
    {
        var chart = BarChart.Create()
            .WithTitle("Sales Chart");

        Assert.Equal("Sales Chart", chart.Title);
    }

    [Fact]
    public void WithTitle_NullTitle_ThrowsArgumentException()
    {
        var chart = BarChart.Create();
        string? nullTitle = null;

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(nullTitle));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithTitle_EmptyTitle_ThrowsArgumentException()
    {
        var chart = BarChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(""));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithTitle_WhitespaceTitle_ThrowsArgumentException()
    {
        var chart = BarChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle("   "));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithDataRange_SetsRanges()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = BarChart.Create()
            .WithDataRange(categoriesRange, valuesRange);

        Assert.Equal(categoriesRange, chart.CategoriesRange);
        Assert.Equal(valuesRange, chart.ValuesRange);
    }

    [Fact]
    public void WithDataRange_SingleCellCategoriesRange_ThrowsArgumentException()
    {
        var categoriesRange = new CellRange(new(0, 0), new(0, 0)); // Invalid: single cell
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = BarChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithDataRange(categoriesRange, valuesRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithDataRange_SingleCellValuesRange_ThrowsArgumentException()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = new CellRange(new(1, 0), new(1, 0)); // Invalid: single cell

        var chart = BarChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithDataRange(categoriesRange, valuesRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithPosition_SetsPosition()
    {
        var chart = BarChart.Create()
            .WithPosition(5, 0, 10, 15);

        Assert.NotNull(chart.Position);
        Assert.Equal(5u, chart.Position.FromColumn);
        Assert.Equal(0u, chart.Position.FromRow);
        Assert.Equal(10u, chart.Position.ToColumn);
        Assert.Equal(15u, chart.Position.ToRow);
    }

    [Fact]
    public void WithPosition_InvalidCoordinates_ThrowsArgumentException()
    {
        var chart = BarChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithPosition(10, 15, 5, 0));
        Assert.Contains("from", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_SetsSize()
    {
        var chart = BarChart.Create()
            .WithSize(8000000, 5000000);

        Assert.Equal(8000000, chart.Size.WidthEmus);
        Assert.Equal(5000000, chart.Size.HeightEmus);
    }

    [Fact]
    public void WithSize_InvalidWidth_ThrowsArgumentException()
    {
        var chart = BarChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(0, 5000000));
        Assert.Contains("width", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_InvalidHeight_ThrowsArgumentException()
    {
        var chart = BarChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(8000000, -1));
        Assert.Contains("height", ex.Message.ToLower());
    }

    [Fact]
    public void WithOrientation_SetsOrientation()
    {
        var chart = BarChart.Create()
            .WithOrientation(BarChartOrientation.Horizontal);

        Assert.Equal(BarChartOrientation.Horizontal, chart.Orientation);
    }

    [Fact]
    public void AddSeries_AddsSeriesAndReturnsBarChart()
    {
        var dataRange = CellRange.FromBounds(1, 0, 1, 5);
        
        var chart = BarChart.Create()
            .AddSeries("Sales", dataRange);

        Assert.Single(chart.Series);
        Assert.Equal("Sales", chart.Series[0].Name);
        Assert.Equal(dataRange, chart.Series[0].DataRange);
    }

    [Fact]
    public void AddSeries_MultipleSeries_AddsAll()
    {
        var dataRange1 = CellRange.FromBounds(1, 0, 1, 5);
        var dataRange2 = CellRange.FromBounds(2, 0, 2, 5);

        var chart = BarChart.Create()
            .AddSeries("Q1", dataRange1)
            .AddSeries("Q2", dataRange2);

        Assert.Equal(2, chart.Series.Count);
        Assert.Equal("Q1", chart.Series[0].Name);
        Assert.Equal("Q2", chart.Series[1].Name);
    }

    [Fact]
    public void FluentAPI_AllMethods_ReturnBarChart()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = BarChart.Create()
            .WithTitle("Sales by Region")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(5, 0, 10, 15)
            .WithSize(6000000, 4000000)
            .WithOrientation(BarChartOrientation.Horizontal)
            .AddSeries("2024", valuesRange);

        Assert.NotNull(chart);
        Assert.Equal("Sales by Region", chart.Title);
        Assert.Equal(categoriesRange, chart.CategoriesRange);
        Assert.Equal(valuesRange, chart.ValuesRange);
        Assert.Equal(5u, chart.Position?.FromColumn);
        Assert.Equal(6000000, chart.Size.WidthEmus);
        Assert.Equal(BarChartOrientation.Horizontal, chart.Orientation);
        Assert.Single(chart.Series);
    }

    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var chart = BarChart.Create();

        Assert.Null(chart.Title);
        Assert.Null(chart.Position);
        Assert.Null(chart.CategoriesRange);
        Assert.Null(chart.ValuesRange);
        Assert.Equal(ChartSize.Default.WidthEmus, chart.Size.WidthEmus);
        Assert.Equal(ChartSize.Default.HeightEmus, chart.Size.HeightEmus);
        Assert.Equal(BarChartOrientation.Vertical, chart.Orientation);
        Assert.Empty(chart.Series);
    }
}
