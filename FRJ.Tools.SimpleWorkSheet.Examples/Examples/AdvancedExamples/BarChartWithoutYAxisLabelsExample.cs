using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class BarChartWithoutYAxisLabelsExample : IExample
{
    public string Name => "Bar Chart Without Y-Axis Labels";
    public string Description => "Demonstrates bar charts with y-axis tick labels explicitly disabled";


    public int ExampleNumber { get; }

    public BarChartWithoutYAxisLabelsExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Product Sales");

        sheet.AddCell(new(0, 0), "Product", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var products = new[] { "Laptop", "Phone", "Tablet", "Monitor", "Keyboard" };
        var sales = new[] { 45000, 62000, 38000, 28000, 15000 };

        for (uint i = 0; i < products.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), products[i], null);
            sheet.AddCell(new(1, i + 1), sales[i], null);
        }

        var categoriesRange = CellRange.FromBounds(0, 1, 0, (uint)products.Length);
        var valuesRange = CellRange.FromBounds(1, 1, 1, (uint)products.Length);

        var barChart = BarChart.Create()
            .WithTitle("Product Sales without Y-Axis Labels")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(0, 7, 8, 22)
            .WithCategoryAxisTitle("Product")
            .WithValueAxisTitle("Sales ($)")
            .WithYAxisLabels(false);

        sheet.AddChart(barChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_BarChartWithoutYAxisLabels.xlsx");
    }
}
