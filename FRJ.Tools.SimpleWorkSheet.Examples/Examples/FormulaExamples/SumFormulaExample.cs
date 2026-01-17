using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class SumFormulaExample : IExample
{
    public string Name => "SUM Formula";
    public string Description => "Using SUM function with ranges";


    public int ExampleNumber { get; }

    public SumFormulaExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("SumFormula");
        
        sheet.AddCell(0, 0, "Sales", configure: cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var sales = new[] { 100, 150, 200, 175, 225, 300, 250 };
        for (uint i = 0; i < sales.Length; i++) sheet.AddCell(0, i + 1, sales[i], null);

        sheet.AddCell(0, 9, "Total", configure: cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 10, new CellFormula("=SUM(A2:A8)"), configure: cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold())
            .WithFormatCode("$#,##0"));
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_SumFormula.xlsx");
    }
}