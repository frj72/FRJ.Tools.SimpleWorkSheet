# FRJ.Tools.SimpleWorkSheet

A modern, fluent C# library for creating Excel documents with a clean builder pattern API.

[![.NET](https://img.shields.io/badge/.NET-9.0-blue)](https://dotnet.microsoft.com/)
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
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

// Create a worksheet
var sheet = new WorkSheet("MySheet");

// Add a simple cell
sheet.AddCell(0, 0, "Hello World");

// Add a styled cell
sheet.AddCell(1, 1, "Styled Text", cell => cell
    .WithColor("FF0000")
    .WithFont(font => font
        .WithSize(14)
        .Bold()
        .Italic()));

// Add a row of data
var headers = new[] { "Name", "Age", "City" }.Select(h => new CellValue(h));
sheet.AddRow(0, 0, headers, cell => cell
    .WithColor("4472C4")
    .WithFont(font => font.Bold().WithColor("FFFFFF")));
```

## API Overview

### Cell Creation

```csharp
// Simple value
sheet.AddCell(1, 1, "Hello");

// With configuration
sheet.AddCell(1, 1, "Hello", cell => cell
    .WithColor("FFFF00")
    .WithFont(font => font.Bold()));

// Complete styling
sheet.AddCell(1, 1, 123.45m, cell => cell
    .WithStyle(style => style
        .WithFillColor("E0E0E0")
        .WithFont(font => font.WithSize(14).Bold())
        .WithFormatCode("0.00")));
```

### Batch Operations

```csharp
// Add entire row
var data = new[] { "A", "B", "C" }.Select(s => new CellValue(s));
sheet.AddRow(row: 1, startColumn: 0, data, cell => cell.WithColor("EFEFEF"));

// Add entire column
var numbers = new[] { 1, 2, 3 }.Select(i => new CellValue(i));
sheet.AddColumn(column: 1, startRow: 0, numbers);

// Update existing cell
sheet.UpdateCell(1, 1, cell => cell
    .WithValue("Updated")
    .WithColor("00FF00"));
```

### Working with Styles

```csharp
// Create reusable style
var headerStyle = CellStyle.Create(
    fillColor: "4472C4",
    font: CellFont.Create(14, "Calibri", "FFFFFF", bold: true),
    borders: null,
    formatCode: null);

// Apply to multiple cells
sheet.AddStyledCell(1, 1, "Header 1", headerStyle);
sheet.AddStyledCell(1, 2, "Header 2", headerStyle);
```

### Import Tracking

```csharp
// Track import metadata
sheet.AddCell(1, 1, "Imported", cell => cell
    .WithMetadata(meta => meta
        .WithSource("csv")
        .WithImportedAt(DateTime.UtcNow)
        .WithOriginalValue("raw_value")
        .AddCustomData("row", 10)));

// Configure import options
var options = ImportOptionsBuilder.Create()
    .WithSourceIdentifier("csv")
    .WithTrimWhitespace(true)
    .WithColumnParser(0, s => new CellValue(int.Parse(s)))
    .Build();
```

## Examples

### Creating a Simple Table

```csharp
var sheet = new WorkSheet("Data");

// Headers
var headers = new[] { "Product", "Price", "Stock" }.Select(h => new CellValue(h));
sheet.AddRow(0, 0, headers, cell => cell
    .WithColor("4472C4")
    .WithFont(font => font.Bold().WithColor("FFFFFF")));

// Data rows
var products = new[]
{
    new[] { "Widget", "9.99", "100" },
    new[] { "Gadget", "19.99", "50" }
};

for (int i = 0; i < products.Length; i++)
{
    var row = products[i].Select(v => new CellValue(v));
    var bgColor = i % 2 == 0 ? "F0F0F0" : "FFFFFF";
    sheet.AddRow((uint)(i + 1), 0, row, cell => cell.WithColor(bgColor));
}
```

### Conditional Formatting

```csharp
var values = new[] { 10, 50, 90 };

foreach (var (value, index) in values.Select((v, i) => (v, i)))
{
    var color = value switch
    {
        < 30 => "FF0000",   // Red
        < 70 => "FFFF00",   // Yellow
        _ => "00FF00"       // Green
    };
    
    sheet.AddCell(index, 0, value, cell => cell
        .WithColor(color)
        .WithFont(font => font.Bold()));
}
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

- .NET 9.0 or later
- Dependencies:
  - DocumentFormat.OpenXml 3.3.0
  - OneOf 3.0.271
  - SkiaSharp 3.119.0


## License

This project is released into the public domain under the [Unlicense](LICENSE).

## Acknowledgments

Built with:
- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) for Excel file generation
- [OneOf](https://github.com/mcintyre321/OneOf) for discriminated unions
- [SkiaSharp](https://github.com/mono/SkiaSharp) for text measurement
