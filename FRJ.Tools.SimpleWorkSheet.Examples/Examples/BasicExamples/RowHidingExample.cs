using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class RowHidingExample : IExample
{
    public string Name => "Row Hiding";
    public string Description => "Demonstrates hiding rows in various ways";

    public void Run()
    {
        var sheet = new WorkSheet("HiddenRows");
        
        sheet.AddCell(new(0, 0), "Visible Rows Example", cell => cell.WithFont(f => f.Bold().WithSize(14)));
        sheet.AddCell(new(0, 1), "Only even rows (2, 4, 6, 8, 10) are visible below", null);
        
        sheet.AddCell(new(0, 3), "Row", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 3), "Data A", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(2, 3), "Data B", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(3, 3), "Data C", cell => cell.WithFont(f => f.Bold()));
        
        for (uint row = 4; row <= 13; row++)
        {
            sheet.AddCell(new(0, row), $"Row {row - 3}", null);
            sheet.AddCell(new(1, row), $"A{row - 3}", null);
            sheet.AddCell(new(2, row), $"B{row - 3}", null);
            sheet.AddCell(new(3, row), $"C{row - 3}", null);
        }
        
        sheet.HideRows(5, 7, 9, 11, 13);
        
        sheet.AddCell(new(0, 15), "Method 2: Using SetRowHidden", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(0, 16), "Row 17 below is hidden using SetRowHidden", null);
        
        sheet.AddCell(new(0, 17), "This row is HIDDEN", null);
        sheet.SetRowHidden(17, true);
        
        sheet.AddCell(new(0, 18), "Row above (17) should be hidden", null);
        
        sheet.AddCell(new(0, 20), "Method 3: Using RowHeight.Hidden enum", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(0, 21), "Row 22 below is hidden using SetRowHeight(22, RowHeight.Hidden)", null);
        
        sheet.AddCell(new(0, 22), "This row is HIDDEN via enum", null);
        sheet.SetRowHeight(22, RowHeight.Hidden);
        
        sheet.AddCell(new(0, 23), "Row above (22) should be hidden", null);
        
        sheet.SetColumnWidth(0, 25.0);
        sheet.SetColumnWidth(1, 15.0);
        sheet.SetColumnWidth(2, 15.0);
        sheet.SetColumnWidth(3, 15.0);
        
        ExampleRunner.SaveWorkSheet(sheet, "87_RowHiding.xlsx");
    }
}
