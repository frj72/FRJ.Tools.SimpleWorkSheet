using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class MinMaxFormulaExample : IExample
{
    public string Name => "MIN/MAX Formulas";
    public string Description => "Finding minimum and maximum values";


    public int ExampleNumber { get; }

    public MinMaxFormulaExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("MinMaxFormula");
        
        sheet.AddCell(0, 0, "Daily Temperatures", configure: cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var temps = new[] { 72, 68, 75, 80, 77, 73, 69 };
        for (uint i = 0; i < temps.Length; i++) sheet.AddCell(0, i + 1, temps[i], null);

        sheet.AddCell(0, 9, "Minimum", configure: cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 10, new CellFormula("=MIN(A2:A8)"), configure: cell => cell
            .WithColor("ADD8E6")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 11, "Maximum", configure: cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 12, new CellFormula("=MAX(A2:A8)"), configure: cell => cell
            .WithColor("FFB6C1")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 13, "Range", configure: cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 14, new CellFormula("=MAX(A2:A8)-MIN(A2:A8)"), configure: cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold()));
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_MinMaxFormula.xlsx");
    }
}