using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartImprovementsTests
{
    [Fact]
    public void BarChart_WithLegendPosition_SetsPosition()
    {
        var chart = BarChart.Create()
            .WithLegendPosition(ChartLegendPosition.Top);
        
        Assert.Equal(ChartLegendPosition.Top, chart.LegendPosition);
    }

    [Fact]
    public void BarChart_WithCategoryAxisTitle_SetsTitle()
    {
        var chart = BarChart.Create()
            .WithCategoryAxisTitle("Products");
        
        Assert.Equal("Products", chart.CategoryAxisTitle);
    }

    [Fact]
    public void BarChart_WithValueAxisTitle_SetsTitle()
    {
        var chart = BarChart.Create()
            .WithValueAxisTitle("Sales ($)");
        
        Assert.Equal("Sales ($)", chart.ValueAxisTitle);
    }

    [Fact]
    public void BarChart_WithDataLabels_SetsDataLabels()
    {
        var chart = BarChart.Create()
            .WithDataLabels(true);
        
        Assert.True(chart.ShowDataLabels);
    }

    [Fact]
    public void BarChart_WithMajorGridlines_SetsGridlines()
    {
        var chart = BarChart.Create()
            .WithMajorGridlines(false);
        
        Assert.False(chart.ShowMajorGridlines);
    }

    [Fact]
    public void Chart_DefaultLegendPosition_IsRight()
    {
        var chart = BarChart.Create();
        
        Assert.Equal(ChartLegendPosition.Right, chart.LegendPosition);
    }

    [Fact]
    public void Chart_DefaultShowMajorGridlines_IsTrue()
    {
        var chart = BarChart.Create();
        
        Assert.True(chart.ShowMajorGridlines);
    }

    [Fact]
    public void Chart_DefaultShowDataLabels_IsFalse()
    {
        var chart = BarChart.Create();
        
        Assert.False(chart.ShowDataLabels);
    }

    [Fact]
    public void LineChart_FluentAPI_AllImprovements()
    {
        var chart = LineChart.Create()
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .WithCategoryAxisTitle("Time")
            .WithValueAxisTitle("Value")
            .WithDataLabels(true)
            .WithMajorGridlines(false);
        
        Assert.Equal(ChartLegendPosition.Bottom, chart.LegendPosition);
        Assert.Equal("Time", chart.CategoryAxisTitle);
        Assert.Equal("Value", chart.ValueAxisTitle);
        Assert.True(chart.ShowDataLabels);
        Assert.False(chart.ShowMajorGridlines);
    }

    [Fact]
    public void PieChart_WithLegendNone_HidesLegend()
    {
        var chart = PieChart.Create()
            .WithLegendPosition(ChartLegendPosition.None);
        
        Assert.Equal(ChartLegendPosition.None, chart.LegendPosition);
    }

    [Fact]
    public void AreaChart_AllImprovements_SetsCorrectly()
    {
        var chart = AreaChart.Create()
            .WithLegendPosition(ChartLegendPosition.Left)
            .WithCategoryAxisTitle("Months")
            .WithValueAxisTitle("Revenue")
            .WithDataLabels(true)
            .WithMajorGridlines(true);
        
        Assert.Equal(ChartLegendPosition.Left, chart.LegendPosition);
        Assert.Equal("Months", chart.CategoryAxisTitle);
        Assert.Equal("Revenue", chart.ValueAxisTitle);
        Assert.True(chart.ShowDataLabels);
        Assert.True(chart.ShowMajorGridlines);
    }

    [Fact]
    public void ScatterChart_WithLegendPosition_SetsPosition()
    {
        var chart = ScatterChart.Create()
            .WithLegendPosition(ChartLegendPosition.Top);
        
        Assert.Equal(ChartLegendPosition.Top, chart.LegendPosition);
    }

    [Fact]
    public void ScatterChart_WithAxisTitles_SetsTitles()
    {
        var chart = ScatterChart.Create()
            .WithCategoryAxisTitle("X Values")
            .WithValueAxisTitle("Y Values");
        
        Assert.Equal("X Values", chart.CategoryAxisTitle);
        Assert.Equal("Y Values", chart.ValueAxisTitle);
    }

    [Fact]
    public void ScatterChart_WithDataLabels_SetsDataLabels()
    {
        var chart = ScatterChart.Create()
            .WithDataLabels(true);
        
        Assert.True(chart.ShowDataLabels);
    }

    [Fact]
    public void ScatterChart_WithMajorGridlines_SetsGridlines()
    {
        var chart = ScatterChart.Create()
            .WithMajorGridlines(false);
        
        Assert.False(chart.ShowMajorGridlines);
    }

    [Fact]
    public void PieChart_WithDataLabels_SetsDataLabels()
    {
        var chart = PieChart.Create()
            .WithDataLabels(true);
        
        Assert.True(chart.ShowDataLabels);
    }

    [Fact]
    public void BarChart_WithImprovements_GeneratesValidFile()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Category", null);
        sheet.AddCell(new(1, 0), "Value", null);
        sheet.AddCell(new(0, 1), "A", null);
        sheet.AddCell(new(1, 1), 100, null);
        sheet.AddCell(new(0, 2), "B", null);
        sheet.AddCell(new(1, 2), 150, null);
        
        var chart = BarChart.Create()
            .WithTitle("Test Chart")
            .WithPosition(3, 0, 10, 15)
            .WithSize(5000000, 3000000)
            .WithLegendPosition(ChartLegendPosition.Top)
            .WithCategoryAxisTitle("Categories")
            .WithValueAxisTitle("Values")
            .WithDataLabels(true)
            .WithMajorGridlines(false)
            .AddSeries("Data", CellRange.FromBounds(1, 1, 1, 2));
        
        sheet.AddChart(chart);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        
        Assert.NotEmpty(binary);
    }

    [Fact]
    public void LineChart_WithNoLegend_GeneratesValidFile()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "X", null);
        sheet.AddCell(new(1, 0), "Y", null);
        sheet.AddCell(new(0, 1), 1, null);
        sheet.AddCell(new(1, 1), 10, null);
        sheet.AddCell(new(0, 2), 2, null);
        sheet.AddCell(new(1, 2), 20, null);
        
        var chart = LineChart.Create()
            .WithTitle("No Legend Chart")
            .WithPosition(3, 0, 10, 15)
            .WithSize(5000000, 3000000)
            .WithLegendPosition(ChartLegendPosition.None)
            .WithCategoryAxisTitle("Time")
            .WithValueAxisTitle("Value")
            .WithMajorGridlines(true)
            .AddSeries("Series 1", CellRange.FromBounds(1, 1, 1, 2));
        
        sheet.AddChart(chart);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        
        Assert.NotEmpty(binary);
    }

    [Fact]
    public void LineChart_WithYAxisLabels_SetsProperty()
    {
        var chart = LineChart.Create()
            .WithYAxisLabels(true);
        
        Assert.True(chart.ShowYAxisLabels);
    }

    [Fact]
    public void LineChart_DefaultYAxisLabels_IsFalse()
    {
        var chart = LineChart.Create();
        
        Assert.False(chart.ShowYAxisLabels);
    }

    [Fact]
    public void AreaChart_WithYAxisLabels_SetsProperty()
    {
        var chart = AreaChart.Create()
            .WithYAxisLabels(true);
        
        Assert.True(chart.ShowYAxisLabels);
    }

    [Fact]
    public void AreaChart_DefaultYAxisLabels_IsFalse()
    {
        var chart = AreaChart.Create();
        
        Assert.False(chart.ShowYAxisLabels);
    }

    [Fact]
    public void ScatterChart_WithYAxisLabels_SetsProperty()
    {
        var chart = ScatterChart.Create()
            .WithYAxisLabels(true);
        
        Assert.True(chart.ShowYAxisLabels);
    }

    [Fact]
    public void ScatterChart_DefaultYAxisLabels_IsFalse()
    {
        var chart = ScatterChart.Create();
        
        Assert.False(chart.ShowYAxisLabels);
    }
}
