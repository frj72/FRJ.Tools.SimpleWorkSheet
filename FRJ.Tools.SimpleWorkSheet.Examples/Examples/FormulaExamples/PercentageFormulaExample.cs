using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;

public class PercentageFormulaExample : IExample
{
    public string Name => "Percentage Formula";
    public string Description => "Calculating percentages and ratios";

    private static readonly string[] SourceArray = ["Product", "Sales", "Target", "Achievement %"];

    public void Run()
    {
        var sheet = new WorkSheet("PercentageFormula");
        
        var headers = SourceArray.Select(h => new CellValue(h));
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