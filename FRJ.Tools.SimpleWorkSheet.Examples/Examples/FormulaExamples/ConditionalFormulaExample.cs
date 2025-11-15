using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class ConditionalFormulaExample : IExample
{
    public string Name => "IF Formula";
    public string Description => "Conditional logic with IF function";

    private static readonly string[] SourceArray = ["Score", "Grade"];

    public void Run()
    {
        var sheet = new WorkSheet("ConditionalFormula");
        
        var headers = SourceArray.Select(h => new CellValue(h));
        sheet.AddRow(0, 0, headers, cell => cell
            .WithColor("2E75B6")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var scores = new[] { 95, 85, 75, 65, 55, 45 };
        
        for (uint i = 0; i < scores.Length; i++)
        {
            var row = i + 1;
            sheet.AddCell(0, row, scores[i]);
            sheet.AddCell(1, row, new CellFormula($"=IF(A{row + 1}>=90,\"A\",IF(A{row + 1}>=80,\"B\",IF(A{row + 1}>=70,\"C\",IF(A{row + 1}>=60,\"D\",\"F\"))))"), cell => cell
                .WithFont(font => font.Bold()));
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "25_ConditionalFormula.xlsx");
    }
}