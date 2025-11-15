using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class CountFormulaExample : IExample
{
    public string Name => "COUNT Formulas";
    public string Description => "Counting cells with COUNT and COUNTA";

    public void Run()
    {
        var sheet = new WorkSheet("CountFormulas");
        
        sheet.AddCell(0, 0, "Values", cell => cell.WithFont(font => font.Bold()));
        var values = new object[] { 10, 20, "Text", 30, "", 40, "More" };
        for (uint i = 0; i < values.Length; i++) sheet.AddCell(0, i + 1, new CellValue(values[i].ToString() ?? ""));

        sheet.AddCell(0, 9, "COUNT (numbers)", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(0, 10, new CellFormula("=COUNT(A2:A8)"), cell => cell
            .WithColor("ADD8E6"));
        
        sheet.AddCell(0, 11, "COUNTA (non-empty)", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(0, 12, new CellFormula("=COUNTA(A2:A8)"), cell => cell
            .WithColor("ADD8E6"));
        
        ExampleRunner.SaveWorkSheet(sheet, "27_CountFormula.xlsx");
    }
}