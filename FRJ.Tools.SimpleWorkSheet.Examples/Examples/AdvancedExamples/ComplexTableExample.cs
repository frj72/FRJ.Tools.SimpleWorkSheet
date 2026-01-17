using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ComplexTableExample : IExample
{
    public string Name => "Complex Table";
    public string Description => "Multi-row, multi-column table with headers and data";


    public int ExampleNumber { get; }

    public ComplexTableExample(int exampleNumber) => ExampleNumber = exampleNumber;
    private static readonly string[] SourceArray = ["ID", "Product", "Category", "Price", "Stock", "Status"];

    public void Run()
    {
        var sheet = new WorkSheet("ComplexTable");
        
        var headers = SourceArray.Select(h => new CellValue(h));
        
        sheet.AddRow(0, 0, headers, configure: cell => cell
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
            
            sheet.AddRow(row, 0, values, configure: cell => cell.WithColor(bgColor));
            
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
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_ComplexTable.xlsx");
    }
}