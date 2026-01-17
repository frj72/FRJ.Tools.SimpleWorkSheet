using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class AreaChartWithoutYAxisLabelsExample : IExample
{
    public string Name => "Area Chart Without Y-Axis Labels";
    public string Description => "Demonstrates area charts with y-axis tick labels explicitly disabled";


    public int ExampleNumber { get; }

    public AreaChartWithoutYAxisLabelsExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Website Traffic");

        sheet.AddCell(new(0, 0), "Month", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Visitors", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        var visitors = new[] { 12000, 15000, 14000, 18000, 21000, 19000 };

        for (uint i = 0; i < months.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), months[i], null);
            sheet.AddCell(new(1, i + 1), visitors[i], null);
        }

        var categoriesRange = CellRange.FromBounds(0, 1, 0, (uint)months.Length);
        var valuesRange = CellRange.FromBounds(1, 1, 1, (uint)months.Length);

        var areaChart = AreaChart.Create()
            .WithTitle("Website Traffic without Y-Axis Labels")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(0, 8, 10, 23)
            .WithCategoryAxisTitle("Month")
            .WithValueAxisTitle("Visitors")
            .WithYAxisLabels(false);

        sheet.AddChart(areaChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_AreaChartWithoutYAxisLabels.xlsx");
    }
}
