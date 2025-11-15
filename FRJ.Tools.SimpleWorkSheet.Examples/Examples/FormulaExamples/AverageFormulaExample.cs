using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class AverageFormulaExample : IExample
{
    public string Name => "AVERAGE Formula";
    public string Description => "Calculating averages with formulas";

    private static readonly string[] SourceArray = ["Student", "Test 1", "Test 2", "Test 3", "Average"];

    public void Run()
    {
        var sheet = new WorkSheet("AverageFormula");
        
        var headers = SourceArray.Select(h => new CellValue(h));
        sheet.AddRow(0, 0, headers, cell => cell
            .WithColor("2E75B6")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var students = new[]
        {
            new { Name = "Alice", Scores = new[] { 85, 90, 88 } },
            new { Name = "Bob", Scores = new[] { 78, 82, 80 } },
            new { Name = "Charlie", Scores = new[] { 92, 95, 93 } }
        };
        
        for (uint i = 0; i < students.Length; i++)
        {
            var row = i + 1;
            var student = students[i];
            
            sheet.AddCell(0, row, student.Name);
            sheet.AddCell(1, row, student.Scores[0]);
            sheet.AddCell(2, row, student.Scores[1]);
            sheet.AddCell(3, row, student.Scores[2]);
            sheet.AddCell(4, row, new CellFormula($"=AVERAGE(B{row + 1}:D{row + 1})"), cell => cell
                .WithFont(font => font.Bold())
                .WithFormatCode("0.0"));
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "23_AverageFormula.xlsx");
    }
}