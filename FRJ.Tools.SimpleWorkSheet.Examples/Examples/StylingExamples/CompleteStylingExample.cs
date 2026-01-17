using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;

public class CompleteStylingExample : IExample
{
    public string Name => "Complete Styling";
    public string Description => "Combining all style elements";


    public int ExampleNumber { get; }

    public CompleteStylingExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("CompleteStyling");
        
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        sheet.AddCell(0, 0, 123.456m, configure: cell => cell
            .WithStyle(style => style
                .WithFillColor("E0E0E0")
                .WithFont(font => font
                    .WithSize(14)
                    .WithName("Calibri")
                    .WithColor("0000FF")
                    .Bold()
                    .Italic())
                .WithBorders(borders)
                .WithFormatCode("0.00")));
        
        sheet.AddCell(0, 1, "Header Style", configure: cell => cell
            .WithColor("4472C4")
            .WithFont(font => font
                .WithSize(16)
                .WithColor("FFFFFF")
                .Bold())
            .WithBorders(borders));
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_CompleteStyling.xlsx");
    }
}