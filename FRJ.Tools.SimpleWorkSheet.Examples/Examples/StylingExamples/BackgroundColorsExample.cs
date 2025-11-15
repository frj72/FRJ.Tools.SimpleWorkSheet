using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;

public class BackgroundColorsExample : IExample
{
    public string Name => "Background Colors";
    public string Description => "Different fill colors";

    public void Run()
    {
        var sheet = new WorkSheet("BackgroundColors");
        
        sheet.AddCell(0, 0, "Red Background", cell => cell.WithColor("FF0000"));
        sheet.AddCell(0, 1, "Green Background", cell => cell.WithColor("00FF00"));
        sheet.AddCell(0, 2, "Blue Background", cell => cell.WithColor("0000FF"));
        sheet.AddCell(0, 3, "Yellow Background", cell => cell.WithColor("FFFF00"));
        sheet.AddCell(0, 4, "Gray Background", cell => cell.WithColor("CCCCCC"));
        sheet.AddCell(0, 5, "Light Blue Background", cell => cell.WithColor("ADD8E6"));
        
        ExampleRunner.SaveWorkSheet(sheet, "07_BackgroundColors.xlsx");
    }
}