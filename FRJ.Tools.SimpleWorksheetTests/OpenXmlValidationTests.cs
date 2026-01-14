using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class OpenXmlValidationTests
{
    private static void ValidateExcelFile(string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        if (errors.Count == 0) return;
        var errorMessages = string.Join("\n", errors.Select(e => 
            $"  - {e.Description}\n    Path: {e.Path?.XPath}\n    Part: {e.Part?.Uri}"));
        Assert.Fail($"OpenXML validation failed:\n{errorMessages}");
    }

    [Fact]
    public void SimplestSheet_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Hello", null);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void RoundTrip_SimplestCase_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Hello", null);
        sheet.AddCell(new(1, 0), 123, null);
        
        var tempPath1 = Path.GetTempFileName() + ".xlsx";
        var tempPath2 = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook1 = new WorkBook("Test", [sheet]);
            workbook1.SaveToFile(tempPath1);
            
            ValidateExcelFile(tempPath1);
            
            var loaded = WorkBookReader.LoadFromFile(tempPath1);
            var loadedSheet = loaded.Sheets.First();
            
            var workbook2 = new WorkBook("Test", [loadedSheet]);
            workbook2.SaveToFile(tempPath2);
            
            ValidateExcelFile(tempPath2);
        }
        finally
        {
            File.Delete(tempPath1);
            File.Delete(tempPath2);
        }
    }

    [Fact]
    public void RoundTrip_WithHyperlinks_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Link", cell => cell
            .WithHyperlink("https://example.com", null));
        
        var tempPath1 = Path.GetTempFileName() + ".xlsx";
        var tempPath2 = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook1 = new WorkBook("Test", [sheet]);
            workbook1.SaveToFile(tempPath1);
            
            ValidateExcelFile(tempPath1);
            
            var loaded = WorkBookReader.LoadFromFile(tempPath1);
            var loadedSheet = loaded.Sheets.First();
            
            var workbook2 = new WorkBook("Test", [loadedSheet]);
            workbook2.SaveToFile(tempPath2);
            
            ValidateExcelFile(tempPath2);
        }
        finally
        {
            File.Delete(tempPath1);
            File.Delete(tempPath2);
        }
    }

    [Fact]
    public void RoundTrip_WithAllFeatures_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.AddCell(new(0, 0), "Header", cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4")
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Center)));
        
        sheet.AddCell(new(0, 1), "Link", cell => cell
            .WithHyperlink("https://example.com", null));
        
        sheet.SetColumnWidth(0, 20.0);
        sheet.SetRowHeight(0, 25.0);
        sheet.FreezePanes(1, 0);
        
        var tempPath1 = Path.GetTempFileName() + ".xlsx";
        var tempPath2 = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook1 = new WorkBook("Test", [sheet]);
            workbook1.SaveToFile(tempPath1);
            
            ValidateExcelFile(tempPath1);
            
            var loaded = WorkBookReader.LoadFromFile(tempPath1);
            var loadedSheet = loaded.Sheets.First();
            
            var workbook2 = new WorkBook("Test", [loadedSheet]);
            workbook2.SaveToFile(tempPath2);
            
            ValidateExcelFile(tempPath2);
        }
        finally
        {
            File.Delete(tempPath1);
            File.Delete(tempPath2);
        }
    }

    [Fact]
    public void BarChart_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Sales");

        sheet.AddCell(new(0, 0), new("Region"), null);
        sheet.AddCell(new(1, 0), new("Sales"), null);
        sheet.AddCell(new(0, 1), new("North"), null);
        sheet.AddCell(new(1, 1), new(100), null);
        sheet.AddCell(new(0, 2), new("South"), null);
        sheet.AddCell(new(1, 2), new(150), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 2);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 2);

        var chart = BarChart.Create()
            .WithTitle("Sales by Region")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(3, 0, 10, 15)
            .WithOrientation(BarChartOrientation.Vertical);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LineChart_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Trends");

        sheet.AddCell(new(0, 0), new("Month"), null);
        sheet.AddCell(new(1, 0), new("Sales"), null);
        sheet.AddCell(new(0, 1), new("Jan"), null);
        sheet.AddCell(new(1, 1), new(100), null);
        sheet.AddCell(new(0, 2), new("Feb"), null);
        sheet.AddCell(new(1, 2), new(120), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 2);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 2);

        var chart = LineChart.Create()
            .WithTitle("Sales Trend")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LineChart_WithDataLabels_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Trends");

        sheet.AddCell(new(0, 0), new("Month"), null);
        sheet.AddCell(new(1, 0), new("Sales"), null);
        sheet.AddCell(new(0, 1), new("Jan"), null);
        sheet.AddCell(new(1, 1), new(100), null);
        sheet.AddCell(new(0, 2), new("Feb"), null);
        sheet.AddCell(new(1, 2), new(120), null);
        sheet.AddCell(new(0, 3), new("Mar"), null);
        sheet.AddCell(new(1, 3), new(140), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 3);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 3);

        var chart = LineChart.Create()
            .WithTitle("Sales Trend with Data Labels")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(3, 0, 10, 15)
            .WithDataLabels(true);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var chartSpace = chartPart.ChartSpace;
            var lineChartElement = chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().FirstOrDefault();
            Assert.NotNull(lineChartElement);
            
            var dataLabels = lineChartElement.Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
            Assert.NotNull(dataLabels);
            
            var showValue = dataLabels.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ShowValue>().FirstOrDefault();
            Assert.NotNull(showValue);
            Assert.True(showValue.Val?.Value);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LineChart_WithMultipleSeries_AndDataLabels_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Regional Sales");

        sheet.AddCell(new(0, 0), new("Quarter"), null);
        sheet.AddCell(new(1, 0), new("North"), null);
        sheet.AddCell(new(2, 0), new("South"), null);
        
        sheet.AddCell(new(0, 1), new("Q1"), null);
        sheet.AddCell(new(1, 1), new(100), null);
        sheet.AddCell(new(2, 1), new(80), null);
        
        sheet.AddCell(new(0, 2), new("Q2"), null);
        sheet.AddCell(new(1, 2), new(120), null);
        sheet.AddCell(new(2, 2), new(90), null);

        var northRange = CellRange.FromBounds(1, 1, 1, 2);
        var southRange = CellRange.FromBounds(2, 1, 2, 2);

        var chart = LineChart.Create()
            .WithTitle("Regional Comparison")
            .WithPosition(4, 0, 11, 15)
            .WithDataLabels(true)
            .AddSeries("North", northRange)
            .AddSeries("South", southRange);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var chartSpace = chartPart.ChartSpace;
            var lineChartElement = chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().FirstOrDefault();
            Assert.NotNull(lineChartElement);
            
            var dataLabels = lineChartElement.Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
            Assert.NotNull(dataLabels);
            
            var showValue = dataLabels.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ShowValue>().FirstOrDefault();
            Assert.NotNull(showValue);
            Assert.True(showValue.Val?.Value);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void PieChart_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Distribution");

        sheet.AddCell(new(0, 0), new("Category"), null);
        sheet.AddCell(new(1, 0), new("Value"), null);
        sheet.AddCell(new(0, 1), new("A"), null);
        sheet.AddCell(new(1, 1), new(30), null);
        sheet.AddCell(new(0, 2), new("B"), null);
        sheet.AddCell(new(1, 2), new(70), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 2);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 2);

        var chart = PieChart.Create()
            .WithTitle("Distribution")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void ScatterChart_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Correlation");

        sheet.AddCell(new(0, 0), new("X"), null);
        sheet.AddCell(new(1, 0), new("Y"), null);
        sheet.AddCell(new(0, 1), new(1), null);
        sheet.AddCell(new(1, 1), new(10), null);
        sheet.AddCell(new(0, 2), new(2), null);
        sheet.AddCell(new(1, 2), new(20), null);

        var xRange = CellRange.FromBounds(0, 1, 0, 2);
        var yRange = CellRange.FromBounds(1, 1, 1, 2);

        var chart = ScatterChart.Create()
            .WithTitle("X vs Y")
            .WithXyData(xRange, yRange)
            .WithPosition(3, 0, 10, 15);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void ChartOnlySheet_NoCells_ValidatesCorrectly()
    {
        var dataSheet = new WorkSheet("Data");
        dataSheet.AddCell(new(0, 0), new("Category"), null);
        dataSheet.AddCell(new(1, 0), new("Value"), null);
        dataSheet.AddCell(new(0, 1), new("A"), null);
        dataSheet.AddCell(new(1, 1), new(100), null);
        dataSheet.AddCell(new(0, 2), new("B"), null);
        dataSheet.AddCell(new(1, 2), new(200), null);

        var chartSheet = new WorkSheet("Dashboard");

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 2);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 2);

        var chart = BarChart.Create()
            .WithTitle("Data from Another Sheet")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(0, 0, 8, 15)
            .WithDataSourceSheet("Data");

        chartSheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [dataSheet, chartSheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void ExcelTable_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Name", null);
        sheet.AddCell(new(1, 0), "Age", null);
        sheet.AddCell(new(0, 1), "Alice", null);
        sheet.AddCell(new(1, 1), 30, null);
        sheet.AddTable("TestTable", 0, 0, 1, 1);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void InsertImage_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Test with Image", null);
        
        var imageData = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        var image = new WorksheetImage(imageData, ImageFormat.Png, new(2, 2), 100, 100);
        sheet.AddImage(image);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void AreaChart_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Month", null);
        sheet.AddCell(new(1, 0), "Sales", null);
        sheet.AddCell(new(0, 1), "Jan", null);
        sheet.AddCell(new(1, 1), 1000, null);
        sheet.AddCell(new(0, 2), "Feb", null);
        sheet.AddCell(new(1, 2), 1500, null);
        
        var chart = AreaChart.Create()
            .WithTitle("Sales")
            .WithDataRange(CellRange.FromBounds(0, 1, 0, 2), CellRange.FromBounds(1, 1, 1, 2))
            .WithPosition(3, 0, 10, 15)
            .WithSize(5000000, 3000000);
        
        sheet.AddChart(chart);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void StackedAreaChart_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Month", null);
        sheet.AddCell(new(1, 0), "Product A", null);
        sheet.AddCell(new(2, 0), "Product B", null);
        sheet.AddCell(new(0, 1), "Jan", null);
        sheet.AddCell(new(1, 1), 1000, null);
        sheet.AddCell(new(2, 1), 800, null);
        sheet.AddCell(new(0, 2), "Feb", null);
        sheet.AddCell(new(1, 2), 1500, null);
        sheet.AddCell(new(2, 2), 900, null);
        
        var chart = AreaChart.Create()
            .WithTitle("Stacked Sales")
            .WithPosition(4, 0, 11, 15)
            .WithSize(5000000, 3000000)
            .WithStacked(true)
            .AddSeries("Product A", CellRange.FromBounds(1, 1, 1, 2))
            .AddSeries("Product B", CellRange.FromBounds(2, 1, 2, 2));
        
        sheet.AddChart(chart);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void ScatterChart_WithDataLabels_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Correlation");
        sheet.AddCell(new(0, 0), "X", null);
        sheet.AddCell(new(1, 0), "Y", null);
        sheet.AddCell(new(0, 1), 1, null);
        sheet.AddCell(new(1, 1), 10, null);
        sheet.AddCell(new(0, 2), 2, null);
        sheet.AddCell(new(1, 2), 20, null);
        
        var chart = ScatterChart.Create()
            .WithTitle("X vs Y with Data Labels")
            .WithXyData(CellRange.FromBounds(0, 1, 0, 2), CellRange.FromBounds(1, 1, 1, 2))
            .WithPosition(3, 0, 10, 15)
            .WithDataLabels(true);
        
        sheet.AddChart(chart);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var chartSpace = chartPart.ChartSpace;
            var scatterChartElement = chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ScatterChart>().FirstOrDefault();
            Assert.NotNull(scatterChartElement);
            
            var dataLabels = scatterChartElement.Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
            Assert.NotNull(dataLabels);
            
            var showValue = dataLabels.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ShowValue>().FirstOrDefault();
            Assert.NotNull(showValue);
            Assert.True(showValue.Val?.Value);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void PieChart_WithDataLabels_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Distribution");
        sheet.AddCell(new(0, 0), "Category", null);
        sheet.AddCell(new(1, 0), "Value", null);
        sheet.AddCell(new(0, 1), "A", null);
        sheet.AddCell(new(1, 1), 30, null);
        sheet.AddCell(new(0, 2), "B", null);
        sheet.AddCell(new(1, 2), 70, null);
        
        var chart = PieChart.Create()
            .WithTitle("Distribution with Data Labels")
            .WithDataRange(CellRange.FromBounds(0, 1, 0, 2), CellRange.FromBounds(1, 1, 1, 2))
            .WithPosition(3, 0, 10, 15)
            .WithDataLabels(true);
        
        sheet.AddChart(chart);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var chartSpace = chartPart.ChartSpace;
            var pieChartElement = chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PieChart>().FirstOrDefault();
            Assert.NotNull(pieChartElement);
            
            var dataLabels = pieChartElement.Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
            Assert.NotNull(dataLabels);
            
            var showValue = dataLabels.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ShowValue>().FirstOrDefault();
            Assert.NotNull(showValue);
            Assert.True(showValue.Val?.Value);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LineChart_WithYAxisLabels_RendersNumberingFormatAndTickPosition()
    {
        var sheet = new WorkSheet("Trends");

        sheet.AddCell(new(0, 0), new("Month"), null);
        sheet.AddCell(new(1, 0), new("Sales"), null);
        sheet.AddCell(new(0, 1), new("Jan"), null);
        sheet.AddCell(new(1, 1), new(100), null);
        sheet.AddCell(new(0, 2), new("Feb"), null);
        sheet.AddCell(new(1, 2), new(120), null);
        sheet.AddCell(new(0, 3), new("Mar"), null);
        sheet.AddCell(new(1, 3), new(140), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 3);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 3);

        var chart = LineChart.Create()
            .WithTitle("Sales Trend with Y-Axis Labels")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(3, 0, 10, 15)
            .WithYAxisLabels(true);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var chartSpace = chartPart.ChartSpace;
            var lineChartElement = chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().FirstOrDefault();
            Assert.NotNull(lineChartElement);
            
            var plotArea = chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault();
            Assert.NotNull(plotArea);
            
            var valueAxis = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().FirstOrDefault();
            Assert.NotNull(valueAxis);
            
            var numberingFormat = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat>().FirstOrDefault();
            Assert.NotNull(numberingFormat);
            Assert.Equal("General", numberingFormat.FormatCode?.Value);
            
            var tickLabelPosition = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.TickLabelPosition>().FirstOrDefault();
            Assert.NotNull(tickLabelPosition);
            Assert.Equal(DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues.NextTo, tickLabelPosition.Val?.Value);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LineChart_WithoutYAxisLabels_DoesNotRenderNumberingFormat()
    {
        var sheet = new WorkSheet("Trends");

        sheet.AddCell(new(0, 0), new("Month"), null);
        sheet.AddCell(new(1, 0), new("Sales"), null);
        sheet.AddCell(new(0, 1), new("Jan"), null);
        sheet.AddCell(new(1, 1), new(100), null);
        sheet.AddCell(new(0, 2), new("Feb"), null);
        sheet.AddCell(new(1, 2), new(120), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 2);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 2);

        var chart = LineChart.Create()
            .WithTitle("Sales Trend without Y-Axis Labels")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(3, 0, 10, 15)
            .WithYAxisLabels(false);

        sheet.AddChart(chart);

        var tempPath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);

            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var plotArea = chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault();
            Assert.NotNull(plotArea);
            
            var valueAxis = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().FirstOrDefault();
            Assert.NotNull(valueAxis);
            
            var numberingFormat = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat>().FirstOrDefault();
            Assert.Null(numberingFormat);
            
            var tickLabelPosition = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.TickLabelPosition>().FirstOrDefault();
            Assert.Null(tickLabelPosition);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void AreaChart_WithYAxisLabels_RendersNumberingFormatAndTickPosition()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), "Month", null);
        sheet.AddCell(new(1, 0), "Sales", null);
        sheet.AddCell(new(0, 1), "Jan", null);
        sheet.AddCell(new(1, 1), 1000, null);
        sheet.AddCell(new(0, 2), "Feb", null);
        sheet.AddCell(new(1, 2), 1500, null);
        
        var chart = AreaChart.Create()
            .WithTitle("Sales")
            .WithDataRange(CellRange.FromBounds(0, 1, 0, 2), CellRange.FromBounds(1, 1, 1, 2))
            .WithPosition(3, 0, 10, 15)
            .WithYAxisLabels(true);
        
        sheet.AddChart(chart);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var plotArea = chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault();
            Assert.NotNull(plotArea);
            
            var valueAxis = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().FirstOrDefault();
            Assert.NotNull(valueAxis);
            
            var numberingFormat = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat>().FirstOrDefault();
            Assert.NotNull(numberingFormat);
            
            var tickLabelPosition = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.TickLabelPosition>().FirstOrDefault();
            Assert.NotNull(tickLabelPosition);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void ScatterChart_WithYAxisLabels_RendersNumberingFormatAndTickPosition()
    {
        var sheet = new WorkSheet("Correlation");
        sheet.AddCell(new(0, 0), "X", null);
        sheet.AddCell(new(1, 0), "Y", null);
        sheet.AddCell(new(0, 1), 1, null);
        sheet.AddCell(new(1, 1), 10, null);
        sheet.AddCell(new(0, 2), 2, null);
        sheet.AddCell(new(1, 2), 20, null);
        
        var chart = ScatterChart.Create()
            .WithTitle("X vs Y")
            .WithXyData(CellRange.FromBounds(0, 1, 0, 2), CellRange.FromBounds(1, 1, 1, 2))
            .WithPosition(3, 0, 10, 15)
            .WithYAxisLabels(true);
        
        sheet.AddChart(chart);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
            
            using var doc = SpreadsheetDocument.Open(tempPath, false);
            var chartPart = doc.WorkbookPart?.WorksheetParts.First().DrawingsPart?.ChartParts.First();
            Assert.NotNull(chartPart);
            
            var plotArea = chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault();
            Assert.NotNull(plotArea);
            
            var valueAxis = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().LastOrDefault();
            Assert.NotNull(valueAxis);
            
            var numberingFormat = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat>().FirstOrDefault();
            Assert.NotNull(numberingFormat);
            
            var tickLabelPosition = valueAxis.Descendants<DocumentFormat.OpenXml.Drawing.Charts.TickLabelPosition>().FirstOrDefault();
            Assert.NotNull(tickLabelPosition);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }
}

