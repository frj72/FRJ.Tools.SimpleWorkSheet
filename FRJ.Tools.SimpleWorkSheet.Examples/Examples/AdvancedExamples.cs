using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples;

public class ConditionalFormattingExample : IExample
{
    public string Name => "Conditional Formatting";
    public string Description => "Dynamic styling based on cell values";

    public void Run()
    {
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
        
        ExampleRunner.SaveWorkSheet(sheet, "16_ConditionalFormatting.xlsx");
    }
}

public class ReusableStylesExample : IExample
{
    public string Name => "Reusable Styles";
    public string Description => "Creating and applying style templates";

    public void Run()
    {
        var sheet = new WorkSheet("ReusableStyles");
        
        var headerStyle = CellStyle.Create(
            fillColor: "4472C4",
            font: CellFont.Create(14, "Calibri", "FFFFFF", bold: true),
            borders: null,
            formatCode: null);
        
        var dataStyle = CellStyle.Create(
            fillColor: "F0F0F0",
            font: CellFont.Create(11, "Calibri", "000000"),
            borders: null,
            formatCode: null);
        
        var highlightStyle = CellStyle.Create(
            fillColor: "FFFF00",
            font: CellFont.Create(11, "Calibri", "000000", bold: true),
            borders: null,
            formatCode: null);
        
        sheet.AddStyledCell(0, 0, "Name", headerStyle);
        sheet.AddStyledCell(1, 0, "Department", headerStyle);
        sheet.AddStyledCell(2, 0, "Salary", headerStyle);
        
        sheet.AddStyledCell(0, 1, "John", dataStyle);
        sheet.AddStyledCell(1, 1, "Engineering", dataStyle);
        sheet.AddStyledCell(2, 1, "$100,000", highlightStyle);
        
        sheet.AddStyledCell(0, 2, "Jane", dataStyle);
        sheet.AddStyledCell(1, 2, "Marketing", dataStyle);
        sheet.AddStyledCell(2, 2, "$90,000", dataStyle);
        
        ExampleRunner.SaveWorkSheet(sheet, "17_ReusableStyles.xlsx");
    }
}

public class ComplexTableExample : IExample
{
    public string Name => "Complex Table";
    public string Description => "Multi-row, multi-column table with headers and data";

    public void Run()
    {
        var sheet = new WorkSheet("ComplexTable");
        
        var headers = new[] { "ID", "Product", "Category", "Price", "Stock", "Status" }
            .Select(h => new CellValue(h));
        
        sheet.AddRow(0, 0, headers, cell => cell
            .WithColor("2E75B6")
            .WithFont(font => font
                .WithSize(12)
                .WithColor("FFFFFF")
                .Bold()));
        
        var products = new[]
        {
            new[] { "001", "Laptop", "Electronics", "$999.99", "15", "In Stock" },
            new[] { "002", "Mouse", "Accessories", "$29.99", "50", "In Stock" },
            new[] { "003", "Keyboard", "Accessories", "$79.99", "3", "Low Stock" },
            new[] { "004", "Monitor", "Electronics", "$299.99", "0", "Out of Stock" },
            new[] { "005", "Webcam", "Electronics", "$89.99", "20", "In Stock" }
        };
        
        for (uint i = 0; i < products.Length; i++)
        {
            var row = i + 1;
            var rowData = products[i];
            var values = rowData.Select(v => new CellValue(v));
            var bgColor = i % 2 == 0 ? "F2F2F2" : "FFFFFF";
            
            sheet.AddRow(row, 0, values, cell => cell.WithColor(bgColor));
            
            var stockValue = int.Parse(rowData[4]);
            var statusColor = stockValue switch
            {
                0 => "FF0000",
                < 10 => "FFAA00",
                _ => "00FF00"
            };
            
            sheet.UpdateCell(5, row, cell => cell
                .WithColor(statusColor)
                .WithFont(font => font.Bold()));
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "18_ComplexTable.xlsx");
    }
}
