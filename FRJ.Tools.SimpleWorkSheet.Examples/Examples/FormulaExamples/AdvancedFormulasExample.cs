using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class AdvancedFormulasExample : IExample
{
    public string Name => "Advanced Formulas";
    public string Description => "Demonstrates FormulaBuilder helper methods";

    public void Run()
    {
        var sheet = new WorkSheet("Advanced Formulas");

        sheet.AddCell(new(0, 0), "Category", cell => cell.WithFont(font => font.Bold()).WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Formula", cell => cell.WithFont(font => font.Bold()).WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "Result", cell => cell.WithFont(font => font.Bold()).WithStyle(style => style.WithFillColor("4472C4")));

        var row = 2u;

        sheet.AddCell(new(0, row), "VLOOKUP", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=VLOOKUP(\"Apple\",F2:G5,2,FALSE)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.VLookup("\"Apple\"", "F2:G5", 2, true), null);
        row++;

        sheet.AddCell(new(0, row), "INDEX/MATCH", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=INDEX(G2:G5,MATCH(\"Banana\",F2:F5,0))", null);
        sheet.AddCell(new(2, row), FormulaBuilder.IndexMatch("G2:G5", "\"Banana\"", "F2:F5"), null);
        row++;

        sheet.AddCell(new(0, row), "IF Statement", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=IF(G4>50,\"High\",\"Low\")", null);
        sheet.AddCell(new(2, row), FormulaBuilder.If("G4>50", "\"High\"", "\"Low\""), null);
        row++;

        sheet.AddCell(new(0, row), "NOT", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=NOT(G3>50)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Not("G3>50"), null);
        row++;

        sheet.AddCell(new(0, row), "AND Logic", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=AND(G3>20,G3<80)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.And("G3>20", "G3<80"), null);
        row++;

        sheet.AddCell(new(0, row), "OR Logic", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=OR(G3<30,G3>70)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Or("G3<30", "G3>70"), null);
        row++;

        sheet.AddCell(new(0, row), "CONCATENATE", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=CONCATENATE(F3,\" - \",G3)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Concatenate("F3", "\" - \"", "G3"), null);
        row++;

        sheet.AddCell(new(0, row), "LEFT", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=LEFT(F2,3)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Left("F2", 3), null);
        row++;

        sheet.AddCell(new(0, row), "UPPER", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=UPPER(F2)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Upper("F2"), null);
        row++;

        sheet.AddCell(new(0, row), "ROUND", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=ROUND(G3/3,2)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Round("G3/3", 2), null);
        row++;

        sheet.AddCell(new(0, row), "ABS", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=ABS(G3-100)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Abs("G3-100"), null);
        row++;

        sheet.AddCell(new(0, row), "POWER", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=POWER(G3,2)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Power("G3", "2"), null);
        row++;

        sheet.AddCell(new(0, row), "SQRT", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=SQRT(G4)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Sqrt("G4"), null);
        row++;

        sheet.AddCell(new(0, row), "TODAY", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=TODAY()", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Today(), null);
        row++;

        sheet.AddCell(new(0, row), "DATE", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=DATE(2025,11,16)", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Date(2025, 11, 16), null);
        row++;

        sheet.AddCell(new(0, row), "YEAR", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "=YEAR(TODAY())", null);
        sheet.AddCell(new(2, row), FormulaBuilder.Year("TODAY()"), null);

        sheet.AddCell(new(5, 1), "Lookup Data", cell => cell.WithFont(font => font.Bold()).WithStyle(style => style.WithFillColor("FFD966")));
        sheet.AddCell(new(6, 1), "Price", cell => cell.WithFont(font => font.Bold()).WithStyle(style => style.WithFillColor("FFD966")));

        sheet.AddCell(new(5, 2), "Apple", null);
        sheet.AddCell(new(6, 2), 45.0, null);

        sheet.AddCell(new(5, 3), "Banana", null);
        sheet.AddCell(new(6, 3), 32.0, null);

        sheet.AddCell(new(5, 4), "Cherry", null);
        sheet.AddCell(new(6, 4), 78.0, null);

        sheet.AddCell(new(5, 5), "Date", null);
        sheet.AddCell(new(6, 5), 23.0, null);

        sheet.SetColumnWidth(0, 15.0);
        sheet.SetColumnWidth(1, 35.0);
        sheet.SetColumnWidth(2, 25.0);
        sheet.SetColumnWidth(5, 12.0);
        sheet.SetColumnWidth(6, 10.0);

        ExampleRunner.SaveWorkSheet(sheet, "40_AdvancedFormulas.xlsx");
    }
}
