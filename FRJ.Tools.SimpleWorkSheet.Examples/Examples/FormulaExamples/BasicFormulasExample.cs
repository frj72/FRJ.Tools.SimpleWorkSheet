using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class BasicFormulasExample : IExample
{
    public string Name => "Basic Formulas";
    public string Description => "Simple arithmetic formulas";

    public void Run()
    {
        var sheet = new WorkSheet("BasicFormulas");
        
        sheet.AddCell(0, 0, "Number 1", null);
        sheet.AddCell(0, 1, 10, null);
        
        sheet.AddCell(0, 2, "Number 2", null);
        sheet.AddCell(0, 3, 20, null);
        
        sheet.AddCell(0, 4, "Sum", null);
        sheet.AddCell(0, 5, new CellFormula("=A2+A4"), configure: cell => cell
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 6, "Product", null);
        sheet.AddCell(0, 7, new CellFormula("=A2*A4"), configure: cell => cell
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 8, "Difference", null);
        sheet.AddCell(0, 9, new CellFormula("=A4-A2"), configure: cell => cell
            .WithFont(font => font.Bold()));
        
        ExampleRunner.SaveWorkSheet(sheet, "21_BasicFormulas.xlsx");
    }
}