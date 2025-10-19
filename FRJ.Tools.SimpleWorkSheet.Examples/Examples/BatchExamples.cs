using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples;

public class RowAdditionExample : IExample
{
    public string Name => "Row Addition";
    public string Description => "Adding multiple cells in a row with consistent styling";

    public void Run()
    {
        var sheet = new WorkSheet("RowAddition");
        
        var headers = new[] { "Product", "Price", "Quantity", "Total" }
            .Select(h => new CellValue(h));
        
        sheet.AddRow(0, 0, headers, cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var row1 = new[] { "Widget", "9.99", "10", "99.90" }
            .Select(v => new CellValue(v));
        
        sheet.AddRow(1, 0, row1, cell => cell.WithColor("F0F0F0"));
        
        var row2 = new[] { "Gadget", "19.99", "5", "99.95" }
            .Select(v => new CellValue(v));
        
        sheet.AddRow(2, 0, row2);
        
        ExampleRunner.SaveWorkSheet(sheet, "10_RowAddition.xlsx");
    }
}

public class ColumnAdditionExample : IExample
{
    public string Name => "Column Addition";
    public string Description => "Adding multiple cells in a column";

    public void Run()
    {
        var sheet = new WorkSheet("ColumnAddition");
        
        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May" }
            .Select(m => new CellValue(m));
        
        sheet.AddColumn(0, 0, months, cell => cell
            .WithFont(font => font.Bold()));
        
        var sales = new[] { 100, 150, 200, 175, 225 }
            .Select(s => new CellValue(s));
        
        sheet.AddColumn(1, 0, sales);
        
        ExampleRunner.SaveWorkSheet(sheet, "11_ColumnAddition.xlsx");
    }
}

public class BulkUpdatesExample : IExample
{
    public string Name => "Bulk Updates";
    public string Description => "Using UpdateCell for modifying existing data";

    public void Run()
    {
        var sheet = new WorkSheet("BulkUpdates");
        
        for (uint i = 0; i < 5; i++) sheet.AddCell(0, i, $"Original {i}");

        for (uint i = 0; i < 5; i++)
            sheet.UpdateCell(0, i, cell => cell
                .WithValue($"Updated {i}")
                .WithColor("FFFF00")
                .WithFont(font => font.Italic()));

        ExampleRunner.SaveWorkSheet(sheet, "12_BulkUpdates.xlsx");
    }
}
