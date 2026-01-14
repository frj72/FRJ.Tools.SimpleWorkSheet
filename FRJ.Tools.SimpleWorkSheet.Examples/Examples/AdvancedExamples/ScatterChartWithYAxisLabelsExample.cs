using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ScatterChartWithYAxisLabelsExample : IExample
{
    public string Name => "Scatter Chart With Explicit Y-Axis Labels";
    public string Description => "Demonstrates scatter charts with y-axis tick labels explicitly enabled";

    public void Run()
    {
        var sheet = new WorkSheet("Age vs Income");

        sheet.AddCell(new(0, 0), "Age", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Income ($)", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var ages = new[] { 22, 25, 28, 32, 35, 38, 42, 45 };
        var incomes = new[] { 35000, 42000, 48000, 58000, 65000, 72000, 82000, 88000 };

        for (uint i = 0; i < ages.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), ages[i], null);
            sheet.AddCell(new(1, i + 1), incomes[i], null);
        }

        var xRange = CellRange.FromBounds(0, 1, 0, (uint)ages.Length);
        var yRange = CellRange.FromBounds(1, 1, 1, (uint)incomes.Length);

        var scatterChart = ScatterChart.Create()
            .WithTitle("Age vs Income with Y-Axis Labels")
            .WithXyData(xRange, yRange)
            .WithPosition(3, 0, 13, 15)
            .WithCategoryAxisTitle("Age")
            .WithValueAxisTitle("Income ($)")
            .WithYAxisLabels(true);

        sheet.AddChart(scatterChart);

        ExampleRunner.SaveWorkSheet(sheet, "099_ScatterChartWithExplicitYAxisLabels.xlsx");
    }
}
