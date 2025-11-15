using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;

public class BorderStylesExample : IExample
{
    public string Name => "Border Styles";
    public string Description => "Various border configurations";

    public void Run()
    {
        var sheet = new WorkSheet("BorderStyles");
        
        var thinBorders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        sheet.AddCell(0, 0, "All Borders", cell => cell.WithBorders(thinBorders));
        
        var topBottomBorders = CellBorders.Create(
            null,
            null,
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        sheet.AddCell(0, 1, "Top/Bottom Only", cell => cell.WithBorders(topBottomBorders));
        
        var leftBorder = CellBorders.Create(
            CellBorder.Create(Colors.Red, CellBorderStyle.Thin),
            null,
            null,
            null);
        
        sheet.AddCell(0, 2, "Left Border Red", cell => cell.WithBorders(leftBorder));
        
        ExampleRunner.SaveWorkSheet(sheet, "08_BorderStyles.xlsx");
    }
}