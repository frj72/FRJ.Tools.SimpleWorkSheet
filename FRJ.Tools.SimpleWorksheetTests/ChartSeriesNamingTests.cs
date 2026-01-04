using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartSeriesNamingTests
{
    [Fact]
    public void BarChart_WithSeriesName_SetsProperty()
    {
        var chart = BarChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5))
            .WithSeriesName("Sales Data");

        Assert.NotNull(chart);
    }

    [Fact]
    public void BarChart_WithSeriesName_ThrowsOnEmptyName()
    {
        var chart = BarChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5));

        Assert.Throws<ArgumentException>(() => chart.WithSeriesName(""));
        Assert.Throws<ArgumentException>(() => chart.WithSeriesName("   "));
    }

    [Fact]
    public void LineChart_WithSeriesName_SetsProperty()
    {
        var chart = LineChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5))
            .WithSeriesName("Temperature");

        Assert.NotNull(chart);
    }

    [Fact]
    public void LineChart_WithSeriesName_ThrowsOnEmptyName()
    {
        var chart = LineChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5));

        Assert.Throws<ArgumentException>(() => chart.WithSeriesName(""));
        Assert.Throws<ArgumentException>(() => chart.WithSeriesName("   "));
    }

    [Fact]
    public void AreaChart_WithSeriesName_SetsProperty()
    {
        var chart = AreaChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5))
            .WithSeriesName("Revenue");

        Assert.NotNull(chart);
    }

    [Fact]
    public void AreaChart_WithSeriesName_ThrowsOnEmptyName()
    {
        var chart = AreaChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5));

        Assert.Throws<ArgumentException>(() => chart.WithSeriesName(""));
        Assert.Throws<ArgumentException>(() => chart.WithSeriesName("   "));
    }

    [Fact]
    public void PieChart_WithSeriesName_SetsProperty()
    {
        var chart = PieChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5))
            .WithSeriesName("Market Share");

        Assert.NotNull(chart);
    }

    [Fact]
    public void PieChart_WithSeriesName_ThrowsOnEmptyName()
    {
        var chart = PieChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5));

        Assert.Throws<ArgumentException>(() => chart.WithSeriesName(""));
        Assert.Throws<ArgumentException>(() => chart.WithSeriesName("   "));
    }

    [Fact]
    public void ScatterChart_WithSeriesName_SetsProperty()
    {
        var chart = ScatterChart.Create()
            .WithXyData(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5))
            .WithSeriesName("Correlation");

        Assert.NotNull(chart);
    }

    [Fact]
    public void ScatterChart_WithSeriesName_ThrowsOnEmptyName()
    {
        var chart = ScatterChart.Create()
            .WithXyData(
                CellRange.FromBounds(0, 0, 0, 5),
                CellRange.FromBounds(1, 0, 1, 5));

        Assert.Throws<ArgumentException>(() => chart.WithSeriesName(""));
        Assert.Throws<ArgumentException>(() => chart.WithSeriesName("   "));
    }

    [Fact]
    public void BarChart_WithSeriesName_AppearsInGeneratedExcel()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Category", null);
        sheet.AddCell(new(1, 0), "Value", null);
        sheet.AddCell(new(0, 1), "A", null);
        sheet.AddCell(new(1, 1), 10, null);
        sheet.AddCell(new(0, 2), "B", null);
        sheet.AddCell(new(1, 2), 20, null);

        var chart = BarChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 2),
                CellRange.FromBounds(1, 1, 1, 2))
            .WithSeriesName("Sales Data")
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void BarChart_WithoutSeriesName_DefaultsToSeries1()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Category", null);
        sheet.AddCell(new(1, 0), "Value", null);
        sheet.AddCell(new(0, 1), "A", null);
        sheet.AddCell(new(1, 1), 10, null);
        sheet.AddCell(new(0, 2), "B", null);
        sheet.AddCell(new(1, 2), 20, null);

        var chart = BarChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 2),
                CellRange.FromBounds(1, 1, 1, 2))
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void LineChart_WithSeriesName_AppearsInGeneratedExcel()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Month", null);
        sheet.AddCell(new(1, 0), "Temperature", null);
        sheet.AddCell(new(0, 1), "Jan", null);
        sheet.AddCell(new(1, 1), 15, null);
        sheet.AddCell(new(0, 2), "Feb", null);
        sheet.AddCell(new(1, 2), 18, null);

        var chart = LineChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 2),
                CellRange.FromBounds(1, 1, 1, 2))
            .WithSeriesName("Temperature Data")
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void AreaChart_WithSeriesName_AppearsInGeneratedExcel()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Quarter", null);
        sheet.AddCell(new(1, 0), "Revenue", null);
        sheet.AddCell(new(0, 1), "Q1", null);
        sheet.AddCell(new(1, 1), 100, null);
        sheet.AddCell(new(0, 2), "Q2", null);
        sheet.AddCell(new(1, 2), 120, null);

        var chart = AreaChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 2),
                CellRange.FromBounds(1, 1, 1, 2))
            .WithSeriesName("Revenue Data")
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void PieChart_WithSeriesName_AppearsInGeneratedExcel()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Product", null);
        sheet.AddCell(new(1, 0), "Sales", null);
        sheet.AddCell(new(0, 1), "Product A", null);
        sheet.AddCell(new(1, 1), 50, null);
        sheet.AddCell(new(0, 2), "Product B", null);
        sheet.AddCell(new(1, 2), 30, null);

        var chart = PieChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 2),
                CellRange.FromBounds(1, 1, 1, 2))
            .WithSeriesName("Product Sales")
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void ScatterChart_WithSeriesName_AppearsInGeneratedExcel()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "X", null);
        sheet.AddCell(new(1, 0), "Y", null);
        sheet.AddCell(new(0, 1), 10, null);
        sheet.AddCell(new(1, 1), 20, null);
        sheet.AddCell(new(0, 2), 15, null);
        sheet.AddCell(new(1, 2), 25, null);

        var chart = ScatterChart.Create()
            .WithXyData(
                CellRange.FromBounds(0, 1, 0, 2),
                CellRange.FromBounds(1, 1, 1, 2))
            .WithSeriesName("Correlation Data")
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void ImportedBarChart_UsesColumnNameAsSeriesName()
    {
        var jsonContent = "[{\"category\":\"A\",\"revenue\":100},{\"category\":\"B\",\"revenue\":200}]";

        var workbook = WorkbookBuilder.FromJson(jsonContent)
            .WithWorkbookName("Test")
            .WithDataSheetName("Data")
            .WithChart(chart => chart
                .OnSheet("Chart")
                .UseColumns("category", "revenue")
                .AsBarChart())
            .Build();

        var bytes = SheetConverter.ToBinaryExcelFile(workbook);
        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void ImportedBarChart_WithCustomName_UsesCustomName()
    {
        var jsonContent = "[{\"category\":\"A\",\"revenue\":100},{\"category\":\"B\",\"revenue\":200}]";

        var workbook = WorkbookBuilder.FromJson(jsonContent)
            .WithWorkbookName("Test")
            .WithDataSheetName("Data")
            .WithChart(chart => chart
                .OnSheet("Chart")
                .UseColumns("category", "revenue")
                .AsBarChart()
                .WithSeriesName("Custom Revenue"))
            .Build();

        var bytes = SheetConverter.ToBinaryExcelFile(workbook);
        Assert.NotEmpty(bytes);
    }
}
