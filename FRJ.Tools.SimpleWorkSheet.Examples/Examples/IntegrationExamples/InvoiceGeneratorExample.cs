using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.IntegrationExamples;

public class InvoiceGeneratorExample : IExample
{
    public string Name => "Invoice Generator";
    public string Description => "Real-world invoice creation";

    private static readonly string[] SourceArray = ["Item", "Description", "Quantity", "Unit Price", "Total"];

    public void Run()
    {
        var sheet = new WorkSheet("Invoice");
        
        sheet.AddCell(0, 0, "INVOICE", cell => cell
            .WithFont(font => font.WithSize(24).Bold())
            .WithColor("4472C4"));
        
        sheet.AddCell(0, 2, "Invoice #:");
        sheet.AddCell(1, 2, "INV-2025-001");
        sheet.AddCell(0, 3, "Date:");
        sheet.AddCell(1, 3, DateTime.Now);
        sheet.AddCell(0, 4, "Due Date:");
        sheet.AddCell(1, 4, DateTime.Now.AddDays(30));
        
        sheet.AddCell(0, 6, "Bill To:");
        sheet.AddCell(0, 7, "John Smith");
        sheet.AddCell(0, 8, "123 Main St");
        sheet.AddCell(0, 9, "New York, NY 10001");
        
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        var headers = SourceArray.Select(h => new CellValue(h));
        
        sheet.AddRow(11, 0, headers, cell => cell
            .WithColor("E7E6E6")
            .WithFont(font => font.Bold())
            .WithBorders(borders));
        
        var items = new[]
        {
            new[] { "001", "Professional Services", "10", "$150.00", "$1,500.00" },
            new[] { "002", "Software License", "1", "$500.00", "$500.00" },
            new[] { "003", "Technical Support", "5", "$100.00", "$500.00" }
        };
        
        for (uint i = 0; i < items.Length; i++)
        {
            var row = i + 12;
            var rowData = items[i].Select(v => new CellValue(v));
            sheet.AddRow(row, 0, rowData, cell => cell.WithBorders(borders));
        }
        
        sheet.AddCell(3, 16, "Subtotal:", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(4, 16, "$2,500.00", cell => cell.WithBorders(borders));
        
        sheet.AddCell(3, 17, "Tax (8%):", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(4, 17, "$200.00", cell => cell.WithBorders(borders));
        
        sheet.AddCell(3, 18, "TOTAL:", cell => cell
            .WithFont(font => font.WithSize(14).Bold())
            .WithColor("4472C4"));
        sheet.AddCell(4, 18, "$2,700.00", cell => cell
            .WithFont(font => font.WithSize(14).Bold())
            .WithColor("4472C4")
            .WithBorders(borders));
        
        ExampleRunner.SaveWorkSheet(sheet, "19_InvoiceGenerator.xlsx");
    }
}