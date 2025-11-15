using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;

public class FontVariationsExample : IExample
{
    public string Name => "Font Variations";
    public string Description => "Bold, italic, size, color, name changes";

    public void Run()
    {
        var sheet = new WorkSheet("FontVariations");
        
        sheet.AddCell(0, 0, "Bold Text", cell => cell
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 1, "Italic Text", cell => cell
            .WithFont(font => font.Italic()));
        
        sheet.AddCell(0, 2, "Large Text", cell => cell
            .WithFont(font => font.WithSize(20)));
        
        sheet.AddCell(0, 3, "Red Text", cell => cell
            .WithFont(font => font.WithColor("FF0000")));
        
        sheet.AddCell(0, 4, "Calibri Font", cell => cell
            .WithFont(font => font.WithName("Calibri")));
        
        sheet.AddCell(0, 5, "Bold Italic Underline", cell => cell
            .WithFont(font => font
                .Bold()
                .Italic()
                .Underline()));
        
        ExampleRunner.SaveWorkSheet(sheet, "06_FontVariations.xlsx");
    }
}