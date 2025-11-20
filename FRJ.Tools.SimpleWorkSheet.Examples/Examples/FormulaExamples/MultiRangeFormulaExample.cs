using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class MultiRangeFormulaExample : IExample
{
    public string Name => "Multi-Range Formulas";
    public string Description => "Using formulas across multiple ranges";

    public void Run()
    {
        var sheet = new WorkSheet("MultiRangeFormula");
        
        sheet.AddCell(0, 0, "Q1 Sales", configure: cell => cell.WithFont(font => font.Bold()));
        var q1Sales = new[] { 1000, 1200, 1100 };
        for (uint i = 0; i < q1Sales.Length; i++) sheet.AddCell(0, i + 1, q1Sales[i], null);

        sheet.AddCell(1, 0, "Q2 Sales", configure: cell => cell.WithFont(font => font.Bold()));
        var q2Sales = new[] { 1300, 1400, 1250 };
        for (uint i = 0; i < q2Sales.Length; i++) sheet.AddCell(1, i + 1, q2Sales[i], null);

        sheet.AddCell(0, 5, "Q1 Total", configure: cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(0, 6, new CellFormula("=SUM(A2:A4)"), configure: cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(1, 5, "Q2 Total", configure: cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(1, 6, new CellFormula("=SUM(B2:B4)"), configure: cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 8, "Grand Total", configure: cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(0, 9, new CellFormula("=SUM(A2:B4)"), configure: cell => cell
            .WithColor("00FF00")
            .WithFont(font => font.Bold())
            .WithFormatCode("$#,##0"));
        
        ExampleRunner.SaveWorkSheet(sheet, "26_MultiRangeFormula.xlsx");
    }
}