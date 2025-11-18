using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ChartFormattingExample : IExample
{
    public string Name => "Chart Formatting";
    public string Description => "Advanced chart formatting: legends, axis titles, data labels, gridlines";

    private static readonly string[] Quarters = ["Q1", "Q2", "Q3", "Q4"];
    private static readonly int[] Revenue = [125000, 142000, 138500, 159200];
    private static readonly int[] Expenses = [95000, 108000, 103000, 119000];

    public void Run()
    {
        var sheet = new WorkSheet("Quarterly Data");

        sheet.AddCell(new(0, 0), "Quarter");
        sheet.AddCell(new(1, 0), "Revenue");
        sheet.AddCell(new(2, 0), "Expenses");

        for (uint i = 0; i < Quarters.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), Quarters[i]);
            sheet.AddCell(new(1, i + 1), Revenue[i]);
            sheet.AddCell(new(2, i + 1), Expenses[i]);
        }

        var chart1 = BarChart.Create()
            .WithTitle("Revenue vs Expenses - Full Formatting")
            .WithPosition(4, 0, 13, 10)
            .WithSize(6000000, 3500000)
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .WithCategoryAxisTitle("Quarter")
            .WithValueAxisTitle("Amount ($)")
            .WithMajorGridlines(true)
            .AddSeries("Revenue", CellRange.FromBounds(1, 1, 1, (uint)Quarters.Length))
            .AddSeries("Expenses", CellRange.FromBounds(2, 1, 2, (uint)Quarters.Length));

        sheet.AddChart(chart1);

        var chart2 = LineChart.Create()
            .WithTitle("Revenue Trend - No Legend")
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, (uint)Quarters.Length),
                CellRange.FromBounds(1, 1, 1, (uint)Quarters.Length))
            .WithPosition(4, 12, 13, 22)
            .WithSize(6000000, 3500000)
            .WithLegendPosition(ChartLegendPosition.None)
            .WithCategoryAxisTitle("Time Period")
            .WithValueAxisTitle("Revenue ($)")
            .WithMajorGridlines(false);

        sheet.AddChart(chart2);

        sheet.AddCell(new(16, 0), "Product");
        sheet.AddCell(new(17, 0), "Sales");
        sheet.AddCell(new(16, 1), "Laptops");
        sheet.AddCell(new(17, 1), 45000);
        sheet.AddCell(new(16, 2), "Phones");
        sheet.AddCell(new(17, 2), 32000);
        sheet.AddCell(new(16, 3), "Tablets");
        sheet.AddCell(new(17, 3), 18000);

        var chart3 = PieChart.Create()
            .WithTitle("Product Distribution")
            .WithDataRange(
                CellRange.FromBounds(16, 1, 16, 3),
                CellRange.FromBounds(17, 1, 17, 3))
            .WithPosition(15, 0, 24, 10)
            .WithSize(5500000, 3500000)
            .WithLegendPosition(ChartLegendPosition.Right)
            .WithDataLabels(true);

        sheet.AddChart(chart3);

        ExampleRunner.SaveWorkSheet(sheet, "53_ChartFormatting.xlsx");
    }
}
