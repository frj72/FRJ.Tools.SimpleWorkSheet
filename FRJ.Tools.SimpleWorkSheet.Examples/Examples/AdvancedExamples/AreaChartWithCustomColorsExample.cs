using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class AreaChartWithCustomColorsExample : IExample
{
    public string Name => "Area Chart With Custom Colors";
    public string Description => "Demonstrates area chart series color customization";


    public int ExampleNumber { get; }

    public AreaChartWithCustomColorsExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Product Revenue");

        sheet.AddCell(new(0, 0), "Month", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Product A", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "Product B", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        var productA = new[] { 40000, 42000, 46000, 48000, 51000, 55000 };
        var productB = new[] { 30000, 34000, 36000, 39000, 41000, 43000 };

        for (uint i = 0; i < months.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), months[i], null);
            sheet.AddCell(new(1, i + 1), productA[i], null);
            sheet.AddCell(new(2, i + 1), productB[i], null);
        }

        var productARange = CellRange.FromBounds(1, 1, 1, (uint)months.Length);
        var productBRange = CellRange.FromBounds(2, 1, 2, (uint)months.Length);

        var areaChart = AreaChart.Create()
            .WithTitle("Product Revenue (Custom Colors)")
            .WithStacked(false)
            .WithPosition(0, 8, 12, 25)
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .AddSeries("Product A", productARange, Colors.LightCobaltBlue)
            .AddSeries("Product B", productBRange, Colors.Mantis);

        sheet.AddChart(areaChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_AreaChartWithCustomColors.xlsx");
    }
}
