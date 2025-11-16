# FRJ.Tools.SimpleWorkSheet

A modern, fluent C# library for creating Excel documents with a clean builder pattern API.

[![.NET](https://img.shields.io/badge/.NET-10.0-blue)](https://dotnet.microsoft.com/)
[![License](https://img.shields.io/badge/license-Unlicense-green)](LICENSE)

## ⚠️ Project Status

This project is currently in **pre-alpha** and will have regularly breaking changes.

It should not be used in any production environment. If some code is usable, fork or copy it and maintain it in a separate/isolated repo.


## Quick Start

### Installation

```bash
   dotnet add package FRJ.Tools.SimpleWorkSheet
```

### Basic Usage

```csharp
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

var sheet = new WorkSheet("HelloWorld");

sheet.AddCell(0, 0, "Hello World");

var workbook = new WorkBook("MyWorkbook", [sheet]);
workbook.SaveToFile("output.xlsx");
```

## API Overview

### Cell Creation

```csharp
sheet.AddCell(0, 0, "String Value");

sheet.AddCell(0, 1, "Integer:");
sheet.AddCell(1, 1, 42);
sheet.AddCell(0, 2, "Decimal:");
sheet.AddCell(1, 2, 3.14159m);
sheet.AddCell(0, 3, "Date:");
sheet.AddCell(1, 3, DateTime.Now);

var borders = CellBorders.Create(
    CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
    CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
    CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
    CellBorder.Create(Colors.Black, CellBorderStyle.Thin));

sheet.AddCell(0, 0, 123.456m, cell => cell
    .WithStyle(style => style
        .WithFillColor("E0E0E0")
        .WithFont(font => font
            .WithSize(14)
            .WithName("Calibri")
            .WithColor("0000FF")
            .Bold()
            .Italic())
        .WithBorders(borders)
        .WithFormatCode("0.00")));
```

### Batch Operations

```csharp
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

var headers = new[] { "Product", "Price", "Quantity", "Total" }
    .Select(h => new CellValue(h));

sheet.AddRow(0, 0, headers, cell => cell
    .WithColor("4472C4")
    .WithFont(font => font.Bold().WithColor("FFFFFF")));

var row1 = new[] { "Widget", "9.99", "10", "99.90" }
    .Select(v => new CellValue(v));

sheet.AddRow(1, 0, row1, cell => cell.WithColor("F0F0F0"));

var numbers = new[] { 1, 2, 3 }.Select(i => new CellValue(i));
sheet.AddColumn(0, 1, numbers);

sheet.UpdateCell(1, 1, cell => cell
    .WithValue("Updated")
    .WithColor("00FF00"));
```

### Working with Styles

```csharp
var headerStyle = CellStyle.Create(
    fillColor: "4472C4",
    font: CellFont.Create(14, "Calibri", "FFFFFF", bold: true),
    borders: null,
    formatCode: null);

sheet.AddStyledCell(1, 1, "Header 1", headerStyle);
sheet.AddStyledCell(1, 2, "Header 2", headerStyle);
```

### Import Tracking

```csharp
sheet.AddCell(1, 1, "Imported", cell => cell
    .WithMetadata(meta => meta
        .WithSource("csv")
        .WithImportedAt(DateTime.UtcNow)
        .WithOriginalValue("raw_value")
        .AddCustomData("row", 10)));

var options = ImportOptionsBuilder.Create()
    .WithSourceIdentifier("csv")
    .WithTrimWhitespace(true)
    .WithColumnParser(0, s => new CellValue(int.Parse(s)))
    .Build();
```

## Examples

### Creating a Simple Table

```csharp
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

var sheet = new WorkSheet("SimpleTable");

sheet.AddCell(0, 0, "Name");
sheet.AddCell(1, 0, "Age");
sheet.AddCell(2, 0, "City");

sheet.AddCell(0, 1, "John");
sheet.AddCell(1, 1, 30);
sheet.AddCell(2, 1, "NYC");

sheet.AddCell(0, 2, "Jane");
sheet.AddCell(1, 2, 25);
sheet.AddCell(2, 2, "LA");
```

### Conditional Formatting

```csharp
var sheet = new WorkSheet("ConditionalFormatting");

sheet.AddCell(0, 0, "Score");
sheet.AddCell(1, 0, "Status");

var scores = new[] { 15, 45, 65, 85, 95 };

for (uint i = 0; i < scores.Length; i++)
{
    var score = scores[i];
    var row = i + 1;
    
    var color = score switch
    {
        < 30 => "FF0000",
        < 70 => "FFFF00",
        _ => "00FF00"
    };
    
    var status = score switch
    {
        < 30 => "Fail",
        < 70 => "Pass",
        _ => "Excellent"
    };
    
    sheet.AddCell(0, row, score, cell => cell
        .WithColor(color)
        .WithFont(font => font.Bold()));
    
    sheet.AddCell(1, row, status, cell => cell
        .WithColor(color));
}
```

### Adding Charts

```csharp
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

var sheet = new WorkSheet("Sales Data");

sheet.AddCell(new(0, 0), "Region", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));
sheet.AddCell(new(1, 0), "Q1 Sales", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));

sheet.AddCell(new(0, 1), new CellValue("North"));
sheet.AddCell(new(1, 1), new CellValue(125000));
sheet.AddCell(new(0, 2), new CellValue("South"));
sheet.AddCell(new(1, 2), new CellValue(98000));
sheet.AddCell(new(0, 3), new CellValue("East"));
sheet.AddCell(new(1, 3), new CellValue(87000));
sheet.AddCell(new(0, 4), new CellValue("West"));
sheet.AddCell(new(1, 4), new CellValue(145000));

var categoriesRange = CellRange.FromBounds(0, 1, 0, 4);
var valuesRange = CellRange.FromBounds(1, 1, 1, 4);

var barChart = BarChart.Create()
    .WithTitle("Q1 Sales by Region")
    .WithDataRange(categoriesRange, valuesRange)
    .WithPosition(0, 6, 8, 21)
    .WithOrientation(BarChartOrientation.Vertical);

sheet.AddChart(barChart);
```

### Chart on Separate Sheet

```csharp
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

var dataSheet = new WorkSheet("Data");

dataSheet.AddCell(new(0, 0), "Month", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));
dataSheet.AddCell(new(1, 0), "Revenue", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));
dataSheet.AddCell(new(2, 0), "Expenses", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));
dataSheet.AddCell(new(3, 0), "Profit", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));

dataSheet.AddCell(new(0, 1), new CellValue("Jan"));
dataSheet.AddCell(new(1, 1), new CellValue(50000));
dataSheet.AddCell(new(2, 1), new CellValue(30000));
dataSheet.AddCell(new(3, 1), new CellValue(20000));

dataSheet.AddCell(new(0, 2), new CellValue("Feb"));
dataSheet.AddCell(new(1, 2), new CellValue(55000));
dataSheet.AddCell(new(2, 2), new CellValue(32000));
dataSheet.AddCell(new(3, 2), new CellValue(23000));

var dashboardSheet = new WorkSheet("Dashboard");

var categoriesRange = CellRange.FromBounds(0, 1, 0, 2);
var revenueRange = CellRange.FromBounds(1, 1, 1, 2);
var profitRange = CellRange.FromBounds(3, 1, 3, 2);

var revenueChart = BarChart.Create()
    .WithTitle("Monthly Revenue")
    .WithDataRange(categoriesRange, revenueRange)
    .WithPosition(0, 0, 8, 15)
    .WithOrientation(BarChartOrientation.Vertical)
    .WithDataSourceSheet("Data");

dashboardSheet.AddChart(revenueChart);

var profitChart = LineChart.Create()
    .WithTitle("Profit Trend")
    .WithDataRange(categoriesRange, profitRange)
    .WithPosition(9, 0, 17, 15)
    .WithMarkers(LineChartMarkerStyle.Circle)
    .WithDataSourceSheet("Data");

dashboardSheet.AddChart(profitChart);

var workbook = new WorkBook("Financial Dashboard", [dataSheet, dashboardSheet]);
workbook.SaveToFile("report.xlsx");
```

### Multiple Chart Types

```csharp
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

var sheet = new WorkSheet("Market Share");

sheet.AddCell(new(0, 0), "Company", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));
sheet.AddCell(new(1, 0), "Market Share %", cell => cell
    .WithFont(font => font.Bold())
    .WithStyle(style => style.WithFillColor("4472C4")));

sheet.AddCell(new(0, 1), new CellValue("Company A"));
sheet.AddCell(new(1, 1), new CellValue(35));

sheet.AddCell(new(0, 2), new CellValue("Company B"));
sheet.AddCell(new(1, 2), new CellValue(25));

sheet.AddCell(new(0, 3), new CellValue("Company C"));
sheet.AddCell(new(1, 3), new CellValue(20));

sheet.AddCell(new(0, 4), new CellValue("Company D"));
sheet.AddCell(new(1, 4), new CellValue(12));

sheet.AddCell(new(0, 5), new CellValue("Others"));
sheet.AddCell(new(1, 5), new CellValue(8));

var categoriesRange = CellRange.FromBounds(0, 1, 0, 5);
var valuesRange = CellRange.FromBounds(1, 1, 1, 5);

var pieChart = PieChart.Create()
    .WithTitle("Market Share Distribution")
    .WithDataRange(categoriesRange, valuesRange)
    .WithPosition(0, 7, 8, 22)
    .WithExplosion(10);

sheet.AddChart(pieChart);

var barChart = BarChart.Create()
    .WithTitle("Market Share Comparison")
    .WithDataRange(categoriesRange, valuesRange)
    .WithPosition(9, 7, 17, 22)
    .WithOrientation(BarChartOrientation.Vertical);

sheet.AddChart(barChart);
```

## Testing

```bash
# Run all tests
dotnet test --verbosity normal

# Run specific test
dotnet test --filter "FullyQualifiedName~CellBuilderTests.CellBuilder_FluentChaining_WorksCorrectly"

# Build release
dotnet build --configuration Release
```

## Requirements

- .NET 10.0 or later
- Dependencies:
  - DocumentFormat.OpenXml 3.3.0
  - OneOf 3.0.271
  - SkiaSharp 3.119.1


## License

This project is released into the public domain under the [Unlicense](LICENSE).

## Acknowledgments

Built with:
- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) for Excel file generation
- [OneOf](https://github.com/mcintyre321/OneOf) for discriminated unions
- [SkiaSharp](https://github.com/mono/SkiaSharp) for text measurement
