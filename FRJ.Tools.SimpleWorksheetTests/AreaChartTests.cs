using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class AreaChartTests
{
    [Fact]
    public void Create_ReturnsAreaChart()
    {
        var chart = AreaChart.Create();
        
        Assert.NotNull(chart);
        Assert.Equal(ChartType.Area, chart.Type);
    }

    [Fact]
    public void WithTitle_ValidTitle_SetsTitle()
    {
        var chart = AreaChart.Create();
        
        chart.WithTitle("Sales Trend");
        
        Assert.Equal("Sales Trend", chart.Title);
    }

    [Fact]
    public void WithTitle_NullTitle_ThrowsException()
    {
        var chart = AreaChart.Create();
        
        string? nullTitle = null;
        var ex = Assert.Throws<ArgumentException>(() => chart.WithTitle(nullTitle));
        
        Assert.Contains("cannot be null or empty", ex.Message);
    }

    [Fact]
    public void WithDataRange_ValidRanges_SetsRanges()
    {
        var chart = AreaChart.Create();
        var categoriesRange = CellRange.FromBounds(0, 0, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 0, 1, 5);
        
        chart.WithDataRange(categoriesRange, valuesRange);
        
        Assert.Equal(categoriesRange, chart.CategoriesRange);
        Assert.Equal(valuesRange, chart.ValuesRange);
    }

    [Fact]
    public void WithPosition_ValidPosition_SetsPosition()
    {
        var chart = AreaChart.Create();
        
        chart.WithPosition(2, 0, 10, 15);
        
        Assert.NotNull(chart.Position);
        Assert.Equal(2u, chart.Position.FromColumn);
        Assert.Equal(0u, chart.Position.FromRow);
        Assert.Equal(10u, chart.Position.ToColumn);
        Assert.Equal(15u, chart.Position.ToRow);
    }

    [Fact]
    public void WithSize_ValidSize_SetsSize()
    {
        var chart = AreaChart.Create();
        
        chart.WithSize(600, 400);
        
        Assert.Equal(600, chart.Size.WidthEmus);
        Assert.Equal(400, chart.Size.HeightEmus);
    }

    [Fact]
    public void WithStacked_True_SetsStacked()
    {
        var chart = AreaChart.Create();
        
        chart.WithStacked(true);
        
        Assert.True(chart.Stacked);
    }

    [Fact]
    public void WithStacked_False_SetsNotStacked()
    {
        var chart = AreaChart.Create();
        
        chart.WithStacked(false);
        
        Assert.False(chart.Stacked);
    }

    [Fact]
    public void Stacked_DefaultValue_IsFalse()
    {
        var chart = AreaChart.Create();
        
        Assert.False(chart.Stacked);
    }

    [Fact]
    public void WithDataSourceSheet_ValidSheetName_SetsDataSourceSheet()
    {
        var chart = AreaChart.Create();
        
        chart.WithDataSourceSheet("DataSheet");
        
        Assert.Equal("DataSheet", chart.DataSourceSheet);
    }

    [Fact]
    public void AddSeries_ValidSeries_AddsSeries()
    {
        var chart = AreaChart.Create();
        var dataRange = CellRange.FromBounds(1, 1, 1, 5);
        
        chart.AddSeries("Series 1", dataRange);
        
        Assert.Single(chart.Series);
        Assert.Equal("Series 1", chart.Series[0].Name);
    }

    [Fact]
    public void FluentAPI_ChainsMethods_ReturnsAreaChart()
    {
        var chart = AreaChart.Create()
            .WithTitle("Area Chart")
            .WithDataRange(CellRange.FromBounds(0, 0, 0, 5), CellRange.FromBounds(1, 0, 1, 5))
            .WithPosition(2, 0, 10, 15)
            .WithSize(500, 300)
            .WithStacked(true)
            .WithDataSourceSheet("Data")
            .AddSeries("Series 1", CellRange.FromBounds(2, 0, 2, 5));
        
        Assert.NotNull(chart);
        Assert.Equal("Area Chart", chart.Title);
        Assert.True(chart.Stacked);
        Assert.Single(chart.Series);
    }

    [Fact]
    public void AreaChart_SaveAndGenerate_CreatesValidFile()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Month");
        sheet.AddCell(new(1, 0), "Sales");
        sheet.AddCell(new(0, 1), "Jan");
        sheet.AddCell(new(1, 1), 1000);
        sheet.AddCell(new(0, 2), "Feb");
        sheet.AddCell(new(1, 2), 1500);
        sheet.AddCell(new(0, 3), "Mar");
        sheet.AddCell(new(1, 3), 1200);
        
        var chart = AreaChart.Create()
            .WithTitle("Monthly Sales")
            .WithDataRange(CellRange.FromBounds(0, 1, 0, 3), CellRange.FromBounds(1, 1, 1, 3))
            .WithPosition(3, 0, 10, 15)
            .WithSize(400, 300);
        
        sheet.AddChart(chart);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        
        Assert.NotEmpty(binary);
    }

    [Fact]
    public void AreaChart_Stacked_SaveAndGenerate_CreatesValidFile()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Month");
        sheet.AddCell(new(1, 0), "Product A");
        sheet.AddCell(new(2, 0), "Product B");
        sheet.AddCell(new(0, 1), "Jan");
        sheet.AddCell(new(1, 1), 1000);
        sheet.AddCell(new(2, 1), 800);
        sheet.AddCell(new(0, 2), "Feb");
        sheet.AddCell(new(1, 2), 1500);
        sheet.AddCell(new(2, 2), 900);
        
        var chart = AreaChart.Create()
            .WithTitle("Stacked Sales")
            .WithPosition(4, 0, 11, 15)
            .WithSize(400, 300)
            .WithStacked(true)
            .AddSeries("Product A", CellRange.FromBounds(1, 1, 1, 2))
            .AddSeries("Product B", CellRange.FromBounds(2, 1, 2, 2));
        
        sheet.AddChart(chart);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        
        Assert.NotEmpty(binary);
    }
}
