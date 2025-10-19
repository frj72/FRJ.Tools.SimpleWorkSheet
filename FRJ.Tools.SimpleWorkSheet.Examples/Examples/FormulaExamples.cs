using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples;

public class BasicFormulasExample : IExample
{
    public string Name => "Basic Formulas";
    public string Description => "Simple arithmetic formulas";

    public void Run()
    {
        var sheet = new WorkSheet("BasicFormulas");
        
        sheet.AddCell(0, 0, "Number 1");
        sheet.AddCell(0, 1, 10);
        
        sheet.AddCell(0, 2, "Number 2");
        sheet.AddCell(0, 3, 20);
        
        sheet.AddCell(0, 4, "Sum");
        sheet.AddCell(0, 5, new CellFormula("=A2+A4"), cell => cell
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 6, "Product");
        sheet.AddCell(0, 7, new CellFormula("=A2*A4"), cell => cell
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 8, "Difference");
        sheet.AddCell(0, 9, new CellFormula("=A4-A2"), cell => cell
            .WithFont(font => font.Bold()));
        
        ExampleRunner.SaveWorkSheet(sheet, "21_BasicFormulas.xlsx");
    }
}

public class SumFormulaExample : IExample
{
    public string Name => "SUM Formula";
    public string Description => "Using SUM function with ranges";

    public void Run()
    {
        var sheet = new WorkSheet("SumFormula");
        
        sheet.AddCell(0, 0, "Sales", cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var sales = new[] { 100, 150, 200, 175, 225, 300, 250 };
        for (uint i = 0; i < sales.Length; i++) sheet.AddCell(0, i + 1, sales[i]);

        sheet.AddCell(0, 9, "Total", cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 10, new CellFormula("=SUM(A2:A8)"), cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold())
            .WithFormatCode("$#,##0"));
        
        ExampleRunner.SaveWorkSheet(sheet, "22_SumFormula.xlsx");
    }
}

public class AverageFormulaExample : IExample
{
    public string Name => "AVERAGE Formula";
    public string Description => "Calculating averages with formulas";

    public void Run()
    {
        var sheet = new WorkSheet("AverageFormula");
        
        var headers = new[] { "Student", "Test 1", "Test 2", "Test 3", "Average" }
            .Select(h => new CellValue(h));
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

public class PercentageFormulaExample : IExample
{
    public string Name => "Percentage Formula";
    public string Description => "Calculating percentages and ratios";

    public void Run()
    {
        var sheet = new WorkSheet("PercentageFormula");
        
        var headers = new[] { "Product", "Sales", "Target", "Achievement %" }
            .Select(h => new CellValue(h));
        sheet.AddRow(0, 0, headers, cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var products = new[]
        {
            new { Name = "Widget", Sales = 8500, Target = 10000 },
            new { Name = "Gadget", Sales = 12000, Target = 10000 },
            new { Name = "Doohickey", Sales = 7500, Target = 8000 }
        };
        
        for (uint i = 0; i < products.Length; i++)
        {
            var row = i + 1;
            var product = products[i];
            
            sheet.AddCell(0, row, product.Name);
            sheet.AddCell(1, row, product.Sales, cell => cell.WithFormatCode("$#,##0"));
            sheet.AddCell(2, row, product.Target, cell => cell.WithFormatCode("$#,##0"));
            sheet.AddCell(3, row, new CellFormula($"=B{row + 1}/C{row + 1}"), cell => cell
                .WithFormatCode("0.0%")
                .WithFont(font => font.Bold()));
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "24_PercentageFormula.xlsx");
    }
}

public class ConditionalFormulaExample : IExample
{
    public string Name => "IF Formula";
    public string Description => "Conditional logic with IF function";

    public void Run()
    {
        var sheet = new WorkSheet("ConditionalFormula");
        
        var headers = new[] { "Score", "Grade" }
            .Select(h => new CellValue(h));
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

public class MultiRangeFormulaExample : IExample
{
    public string Name => "Multi-Range Formulas";
    public string Description => "Using formulas across multiple ranges";

    public void Run()
    {
        var sheet = new WorkSheet("MultiRangeFormula");
        
        sheet.AddCell(0, 0, "Q1 Sales", cell => cell.WithFont(font => font.Bold()));
        var q1Sales = new[] { 1000, 1200, 1100 };
        for (uint i = 0; i < q1Sales.Length; i++) sheet.AddCell(0, i + 1, q1Sales[i]);

        sheet.AddCell(1, 0, "Q2 Sales", cell => cell.WithFont(font => font.Bold()));
        var q2Sales = new[] { 1300, 1400, 1250 };
        for (uint i = 0; i < q2Sales.Length; i++) sheet.AddCell(1, i + 1, q2Sales[i]);

        sheet.AddCell(0, 5, "Q1 Total", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(0, 6, new CellFormula("=SUM(A2:A4)"), cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(1, 5, "Q2 Total", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(1, 6, new CellFormula("=SUM(B2:B4)"), cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 8, "Grand Total", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(0, 9, new CellFormula("=SUM(A2:B4)"), cell => cell
            .WithColor("00FF00")
            .WithFont(font => font.Bold())
            .WithFormatCode("$#,##0"));
        
        ExampleRunner.SaveWorkSheet(sheet, "26_MultiRangeFormula.xlsx");
    }
}

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

public class MinMaxFormulaExample : IExample
{
    public string Name => "MIN/MAX Formulas";
    public string Description => "Finding minimum and maximum values";

    public void Run()
    {
        var sheet = new WorkSheet("MinMaxFormula");
        
        sheet.AddCell(0, 0, "Daily Temperatures", cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var temps = new[] { 72, 68, 75, 80, 77, 73, 69 };
        for (uint i = 0; i < temps.Length; i++) sheet.AddCell(0, i + 1, temps[i]);

        sheet.AddCell(0, 9, "Minimum", cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 10, new CellFormula("=MIN(A2:A8)"), cell => cell
            .WithColor("ADD8E6")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 11, "Maximum", cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 12, new CellFormula("=MAX(A2:A8)"), cell => cell
            .WithColor("FFB6C1")
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 13, "Range", cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(0, 14, new CellFormula("=MAX(A2:A8)-MIN(A2:A8)"), cell => cell
            .WithColor("FFFF00")
            .WithFont(font => font.Bold()));
        
        ExampleRunner.SaveWorkSheet(sheet, "28_MinMaxFormula.xlsx");
    }
}
