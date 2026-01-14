using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ScatterChartWithDataLabelsExample : IExample
{
    public string Name => "Scatter Chart With Data Labels";
    public string Description => "Demonstrates scatter charts with data labels showing y-values";

    public void Run()
    {
        var sheet = new WorkSheet("Temperature vs Sales");

        sheet.AddCell(new(0, 0), "Temperature (F)", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Sales ($)", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var temperatures = new[] { 45, 52, 58, 65, 72, 78, 85, 92 };
        var sales = new[] { 2200, 2500, 2800, 3100, 3500, 3900, 4200, 4500 };

        for (uint i = 0; i < temperatures.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), temperatures[i], null);
            sheet.AddCell(new(1, i + 1), sales[i], null);
        }

        var xRange = CellRange.FromBounds(0, 1, 0, (uint)temperatures.Length);
        var yRange = CellRange.FromBounds(1, 1, 1, (uint)sales.Length);

        var scatterChart = ScatterChart.Create()
            .WithTitle("Temperature vs Sales Correlation")
            .WithXyData(xRange, yRange)
            .WithPosition(3, 0, 13, 15)
            .WithCategoryAxisTitle("Temperature (F)")
            .WithValueAxisTitle("Sales ($)")
            .WithDataLabels(true);

        sheet.AddChart(scatterChart);

        ExampleRunner.SaveWorkSheet(sheet, "096_ScatterChartWithDataLabels.xlsx");
    }
}
