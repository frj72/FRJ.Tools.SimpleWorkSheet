using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class CellPositioningExample : IExample
{
    public string Name => "Cell Positioning";
    public string Description => "Using coordinates vs CellPosition objects";

    public void Run()
    {
        var sheet = new WorkSheet("Positioning");
        
        sheet.AddCell(0, 0, "Using x, y coordinates");
        
        var position = new CellPosition(0, 1);
        sheet.AddCell(position, "Using CellPosition object");
        
        sheet.AddCell(new(0, 2), "Another CellPosition");
        
        ExampleRunner.SaveWorkSheet(sheet, "04_CellPositioning.xlsx");
    }
}