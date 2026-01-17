using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ConditionalHidingExample : IExample
{
    public string Name => "Conditional Hiding";
    public string Description => "Demonstrates hiding rows and columns based on conditions";


    public int ExampleNumber { get; }

    public ConditionalHidingExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("ConditionalHiding");
        
        sheet.AddCell(new(0, 0), "Sales Report - Q1 2025", cell => cell
            .WithFont(f => f.Bold().WithSize(16))
            .WithColor("ADD8E6"));
        sheet.MergeCells(0, 0, 6, 0);
        
        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        sheet.AddCell(new(0, 2), "Product", cell => cell.WithFont(f => f.Bold()));
        for (uint i = 0; i < months.Length; i++)
            sheet.AddCell(new(i + 1, 2), months[i], cell => cell.WithFont(f => f.Bold()));
        
        (string, int[])[] products =
        [
            ("Laptop", [1200, 0, 1500, 0, 1800, 0]),
            ("Mouse", [150, 200, 180, 0, 0, 210]),
            ("Keyboard", [300, 320, 0, 340, 350, 0]),
            ("Monitor", [0, 0, 0, 0, 0, 0]),
            ("Headset", [400, 0, 500, 600, 0, 700])
        ];
        
        var columnsWithSales = new bool[6];
        var rowsWithSales = new List<uint>();
        
        for (uint row = 0; row < products.Length; row++)
        {
            var (product, sales) = products[row];
            var rowIndex = row + 3;
            
            sheet.AddCell(new(0, rowIndex), product, null);
            
            var hasAnySales = false;
            for (uint col = 0; col < sales.Length; col++)
            {
                var value = sales[col];
                sheet.AddCell(new(col + 1, rowIndex), value, cell => cell
                    .WithFormatCode("$#,##0")
                    .WithColor(value == 0 ? "D3D3D3" : "FFFFFF"));
                
                if (value <= 0) continue;
                columnsWithSales[col] = true;
                hasAnySales = true;
            }
            
            if (hasAnySales)
                rowsWithSales.Add(rowIndex);
        }
        
        for (uint col = 0; col < columnsWithSales.Length; col++)
            if (!columnsWithSales[col])
                sheet.HideColumn(col + 1);
        
        for (uint row = 3; row < 3 + products.Length; row++)
            if (!rowsWithSales.Contains(row))
                sheet.HideRow(row);
        
        sheet.AddCell(new(0, 10), "Note: Months with no sales and products with no sales are hidden", cell => cell
            .WithFont(f => f.Italic())
            .WithColor(Colors.Yellow));
        sheet.MergeCells(0, 10, 6, 10);
        
        sheet.AddCell(new(0, 12), "Hidden columns:", cell => cell.WithFont(f => f.Bold()));
        var hiddenCols = new List<string>();
        for (uint i = 0; i < months.Length; i++)
            if (!columnsWithSales[i])
                hiddenCols.Add(months[i]);
        sheet.AddCell(new(1, 12), string.Join(", ", hiddenCols), null);
        
        sheet.AddCell(new(0, 13), "Hidden products:", cell => cell.WithFont(f => f.Bold()));
        var hiddenProducts = new List<string>();
        for (uint i = 0; i < products.Length; i++)
            if (!rowsWithSales.Contains(i + 3))
                hiddenProducts.Add(products[i].Item1);
        sheet.AddCell(new(1, 13), string.Join(", ", hiddenProducts), null);
        
        sheet.SetColumnWidth(0, 20.0);
        for (uint i = 1; i <= 6; i++)
            sheet.SetColumnWidth(i, 12.0);
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_ConditionalHiding.xlsx");
    }
}
