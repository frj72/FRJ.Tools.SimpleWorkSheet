using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class PieChartExample : IExample
{
    public string Name => "Pie Chart";
    public string Description => "Demonstrates pie charts for distribution visualization";


    public int ExampleNumber { get; }

    public PieChartExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Market Share");

        sheet.AddCell(new(0, 0), "Company", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Market Share %", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.AddCell(new(0, 1), new("Company A"), null);
        sheet.AddCell(new(1, 1), new(35), null);

        sheet.AddCell(new(0, 2), new("Company B"), null);
        sheet.AddCell(new(1, 2), new(25), null);

        sheet.AddCell(new(0, 3), new("Company C"), null);
        sheet.AddCell(new(1, 3), new(20), null);

        sheet.AddCell(new(0, 4), new("Company D"), null);
        sheet.AddCell(new(1, 4), new(12), null);

        sheet.AddCell(new(0, 5), new("Others"), null);
        sheet.AddCell(new(1, 5), new(8), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 5);

        var pieChart = PieChart.Create()
            .WithTitle("Market Share Distribution")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(0, 7, 8, 22)
            .WithExplosion(10);

        sheet.AddChart(pieChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_PieChart.xlsx");
    }
}
