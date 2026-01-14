using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class RowHeightExample : IExample
{
    public string Name => "Row Height Control";
    public string Description => "Demonstrates explicit row height settings";

    public void Run()
    {
        var sheet = new WorkSheet("RowHeights");
        
        sheet.AddCell(new(0, 0), "Default Row Height", null);
        sheet.AddCell(new(1, 0), "Row 0 uses default height", null);
        
        sheet.AddCell(new(0, 1), "Small Row", null);
        sheet.AddCell(new(1, 1), "Row 1 has height 15", null);
        sheet.SetRowHeight(1, 15.0);
        
        sheet.AddCell(new(0, 2), "Medium Row", null);
        sheet.AddCell(new(1, 2), "Row 2 has height 30", null);
        sheet.SetRowHeight(2, 30.0);
        
        sheet.AddCell(new(0, 3), "Large Row", null);
        sheet.AddCell(new(1, 3), "Row 3 has height 50", null);
        sheet.SetRowHeight(3, 50.0);
        
        sheet.AddCell(new(0, 4), "Extra Large Row", null);
        sheet.AddCell(new(1, 4), "Row 4 has height 75", null);
        sheet.SetRowHeight(4, 75.0);
        
        sheet.AddCell(new(0, 5), "Styled Content", null);
        sheet.AddCell(new(1, 5), "This row has custom height and styling", cell => cell
            .WithFont(font => font.WithSize(18).Bold())
            .WithColor("E7E6E6"));
        sheet.SetRowHeight(5, 40.0);
        
        ExampleRunner.SaveWorkSheet(sheet, "031_RowHeight.xlsx");
    }
}
