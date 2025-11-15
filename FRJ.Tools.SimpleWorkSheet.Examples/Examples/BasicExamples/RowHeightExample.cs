using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class RowHeightExample : IExample
{
    public string Name => "Row Height Control";
    public string Description => "Demonstrates explicit row height settings";

    public void Run()
    {
        var sheet = new WorkSheet("RowHeights");
        
        sheet.AddCell(new CellPosition(0, 0), "Default Row Height");
        sheet.AddCell(new CellPosition(1, 0), "Row 0 uses default height");
        
        sheet.AddCell(new CellPosition(0, 1), "Small Row");
        sheet.AddCell(new CellPosition(1, 1), "Row 1 has height 15");
        sheet.SetRowHeight(1, 15.0);
        
        sheet.AddCell(new CellPosition(0, 2), "Medium Row");
        sheet.AddCell(new CellPosition(1, 2), "Row 2 has height 30");
        sheet.SetRowHeight(2, 30.0);
        
        sheet.AddCell(new CellPosition(0, 3), "Large Row");
        sheet.AddCell(new CellPosition(1, 3), "Row 3 has height 50");
        sheet.SetRowHeight(3, 50.0);
        
        sheet.AddCell(new CellPosition(0, 4), "Extra Large Row");
        sheet.AddCell(new CellPosition(1, 4), "Row 4 has height 75");
        sheet.SetRowHeight(4, 75.0);
        
        sheet.AddCell(new CellPosition(0, 5), "Styled Content");
        sheet.AddCell(new CellPosition(1, 5), "This row has custom height and styling", cell => cell
            .WithFont(font => font.WithSize(18).Bold())
            .WithColor("E7E6E6"));
        sheet.SetRowHeight(5, 40.0);
        
        ExampleRunner.SaveWorkSheet(sheet, "31_RowHeight.xlsx");
    }
}
