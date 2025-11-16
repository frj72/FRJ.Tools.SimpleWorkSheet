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
    public void WithTitle_EmptyTitle_ThrowsArgumentException()
    {
        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(""));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithTitle_WhitespaceTitle_ThrowsArgumentException()
    {
        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle("   "));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void WithDataRange_SingleCellCategoriesRange_ThrowsArgumentException()
    {
        var categoriesRange = new CellRange(new(0, 0), new(0, 0));
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);

        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithDataRange(categoriesRange, valuesRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithDataRange_SingleCellValuesRange_ThrowsArgumentException()
    {
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = new CellRange(new(1, 0), new(1, 0));

        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithDataRange(categoriesRange, valuesRange));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void WithPosition_InvalidCoordinates_ThrowsArgumentException()
    {
        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithPosition(10, 15, 5, 0));
        Assert.Contains("from", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_SetsSize()
    {
        var chart = PieChart.Create()
            .WithSize(8000000, 5000000);

        Assert.Equal(8000000, chart.Size.WidthEmus);
        Assert.Equal(5000000, chart.Size.HeightEmus);
    }

    [Fact]
    public void WithSize_InvalidWidth_ThrowsArgumentException()
    {
        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(0, 5000000));
        Assert.Contains("width", ex.Message.ToLower());
    }

    [Fact]
    public void WithSize_InvalidHeight_ThrowsArgumentException()
    {
        var chart = PieChart.Create();

        var ex = Assert.Throws<ArgumentException>(() => chart.WithSize(8000000, -1));
        Assert.Contains("height", ex.Message.ToLower());
    }

    [Fact]
    public void AddSeries_AddsSeriesAndReturnsPieChart()
    {
        var dataRange = CellRange.FromBounds(1, 0, 1, 5);
        
        var chart = PieChart.Create()
            .AddSeries("Market", dataRange);

        Assert.Single(chart.Series);
        Assert.Equal("Market", chart.Series[0].Name);
        Assert.Equal(dataRange, chart.Series[0].DataRange);
    }

    [Fact]
    public void AddSeries_MultipleSeries_AddsAll()
    {
        var dataRange1 = CellRange.FromBounds(1, 0, 1, 5);
        var dataRange2 = CellRange.FromBounds(2, 0, 2, 5);

        var chart = PieChart.Create()
            .AddSeries("2023", dataRange1)
            .AddSeries("2024", dataRange2);

        Assert.Equal(2, chart.Series.Count);
        Assert.Equal("2023", chart.Series[0].Name);
        Assert.Equal("2024", chart.Series[1].Name);
    }

    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var chart = PieChart.Create();

        Assert.Null(chart.Title);
        Assert.Null(chart.Position);
        Assert.Null(chart.CategoriesRange);
        Assert.Null(chart.ValuesRange);
        Assert.Equal(ChartSize.Default.WidthEmus, chart.Size.WidthEmus);
        Assert.Equal(ChartSize.Default.HeightEmus, chart.Size.HeightEmus);
        Assert.Equal(0u, chart.Explosion);
        Assert.Equal(0u, chart.FirstSliceAngle);
        Assert.Empty(chart.Series);
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
            .WithSize(6000000, 4000000)
            .WithExplosion(10)
            .WithFirstSliceAngle(45);

        Assert.NotNull(chart);
        Assert.Equal("Market Share", chart.Title);
        Assert.Equal(6000000, chart.Size.WidthEmus);
        Assert.Equal(10u, chart.Explosion);
        Assert.Equal(45u, chart.FirstSliceAngle);
    }
}
