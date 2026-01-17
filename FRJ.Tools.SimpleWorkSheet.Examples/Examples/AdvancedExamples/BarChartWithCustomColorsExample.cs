using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class BarChartWithCustomColorsExample : IExample
{
    public string Name => "Bar Chart With Custom Colors";
    public string Description => "Demonstrates bar chart series color customization";


    public int ExampleNumber { get; }

    public BarChartWithCustomColorsExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Quarterly Sales");

        sheet.AddCell(new(0, 0), "Quarter", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Revenue", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "Expenses", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var quarters = new[] { "Q1", "Q2", "Q3", "Q4" };
        var revenue = new[] { 100000, 120000, 115000, 130000 };
        var expenses = new[] { 65000, 70000, 68000, 75000 };

        for (uint i = 0; i < quarters.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), quarters[i], null);
            sheet.AddCell(new(1, i + 1), revenue[i], null);
            sheet.AddCell(new(2, i + 1), expenses[i], null);
        }

        var revenueRange = CellRange.FromBounds(1, 1, 1, (uint)quarters.Length);
        var expensesRange = CellRange.FromBounds(2, 1, 2, (uint)quarters.Length);

        var barChart = BarChart.Create()
            .WithTitle("Revenue vs Expenses (Custom Colors)")
            .WithPosition(0, 7, 10, 22)
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .WithCategoryAxisTitle("Quarter")
            .WithValueAxisTitle("Amount ($)")
            .AddSeries("Revenue", revenueRange, Colors.Green)
            .AddSeries("Expenses", expensesRange, Colors.Red);

        sheet.AddChart(barChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_BarChartWithCustomColors.xlsx");
    }
}
