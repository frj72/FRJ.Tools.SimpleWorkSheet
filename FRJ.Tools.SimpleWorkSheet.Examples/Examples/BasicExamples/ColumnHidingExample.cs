using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class ColumnHidingExample : IExample
{
    public string Name => "Column Hiding";
    public string Description => "Demonstrates hiding columns in various ways";

    public void Run()
    {
        var sheet = new WorkSheet("HiddenColumns");
        
        sheet.AddCell(new(0, 0), "Column A - Visible", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 0), "Column B - Hidden", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(2, 0), "Column C - Visible", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(3, 0), "Column D - Hidden", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(4, 0), "Column E - Visible", cell => cell.WithFont(f => f.Bold()));
        
        for (uint row = 1; row <= 10; row++)
        {
            sheet.AddCell(new(0, row), $"A{row}", null);
            sheet.AddCell(new(1, row), $"B{row}", null);
            sheet.AddCell(new(2, row), $"C{row}", null);
            sheet.AddCell(new(3, row), $"D{row}", null);
            sheet.AddCell(new(4, row), $"E{row}", null);
        }
        
        sheet.HideColumn(1);
        sheet.HideColumn(3);
        
        sheet.AddCell(new(0, 12), "Method 2: Using SetColumnHidden", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(0, 13), "Column F below is hidden using SetColumnHidden(5, true)", null);
        
        sheet.AddCell(new(5, 0), "Column F - Hidden via SetColumnHidden", cell => cell.WithFont(f => f.Bold()));
        for (uint row = 1; row <= 10; row++)
            sheet.AddCell(new(5, row), $"F{row}", null);
        
        sheet.SetColumnHidden(5, true);
        
        sheet.AddCell(new(0, 15), "Method 3: Using CellWidth.Hidden enum", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(0, 16), "Column G below is hidden using SetColumnWidth(6, CellWidth.Hidden)", null);
        
        sheet.AddCell(new(6, 0), "Column G - Hidden via enum", cell => cell.WithFont(f => f.Bold()));
        for (uint row = 1; row <= 10; row++)
            sheet.AddCell(new(6, row), $"G{row}", null);
        
        sheet.SetColumnWidth(6, CellWidth.Hidden);
        
        sheet.SetColumnWidth(0, 30.0);
        sheet.SetColumnWidth(2, 20.0);
        sheet.SetColumnWidth(4, 20.0);
        
        ExampleRunner.SaveWorkSheet(sheet, "86_ColumnHiding.xlsx");
    }
}
