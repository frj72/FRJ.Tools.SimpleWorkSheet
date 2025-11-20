using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.StressTests;

public class ComplexFormulasExample : IShowcase
{
    public string Name => "Complex Formulas Across Multiple Sheets";
    public string Description => "Tests complex formula references across sheets";
    public string Category => "Stress Tests";

    public void Run()
    {
        var sheet1 = new WorkSheet("Data1");
        var sheet2 = new WorkSheet("Data2");
        var summary = new WorkSheet("Summary");
        
        for (uint row = 0; row < 100; row++)
        {
            sheet1.AddCell(new(0, row), (row + 1) * 10m, null);
            sheet2.AddCell(new(0, row), (row + 1) * 5m, null);
        }
        
        summary.AddCell(new(0, 0), "Calculation", cell => cell.WithFont(f => f.Bold()));
        summary.AddCell(new(1, 0), "Result", cell => cell.WithFont(f => f.Bold()));
        
        summary.AddCell(new(0, 1), "Sum from Sheet1", null);
        summary.AddCell(new(1, 1), new CellFormula("=SUM(Data1!A1:A100)"), null);
        
        summary.AddCell(new(0, 2), "Sum from Sheet2", null);
        summary.AddCell(new(1, 2), new CellFormula("=SUM(Data2!A1:A100)"), null);
        
        summary.AddCell(new(0, 3), "Total from both", null);
        summary.AddCell(new(1, 3), new CellFormula("=B2+B3"), null);
        
        summary.AddCell(new(0, 4), "Average Sheet1", null);
        summary.AddCell(new(1, 4), new CellFormula("=AVERAGE(Data1!A1:A100)"), null);
        
        summary.AddCell(new(0, 5), "Max from Sheet2", null);
        summary.AddCell(new(1, 5), new CellFormula("=MAX(Data2!A1:A100)"), null);
        
        summary.SetColumnWidth(0, 20.0);
        summary.SetColumnWidth(1, 15.0);
        
        var workbook = new WorkBook("ComplexFormulas", [sheet1, sheet2, summary]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_11_ComplexFormulas.xlsx");
    }
}
