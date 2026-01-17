using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ChartSeriesNamesExample : IExample
{
    public string Name => "Chart Series Names";
    public string Description => "Demonstrates custom series names in chart legends";


    public int ExampleNumber { get; }

    public ChartSeriesNamesExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Sales Data");

        sheet.AddCell(new(0, 0), "Month", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Revenue", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "Expenses", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        var revenue = new[] { 45000, 48000, 52000, 49000, 54000, 58000 };
        var expenses = new[] { 32000, 35000, 37000, 34000, 38000, 41000 };

        for (uint i = 0; i < months.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), months[i], null);
            sheet.AddCell(new(1, i + 1), revenue[i], null);
            sheet.AddCell(new(2, i + 1), expenses[i], null);
        }

        var categoriesRange = CellRange.FromBounds(0, 1, 0, (uint)months.Length);
        var revenueRange = CellRange.FromBounds(1, 1, 1, (uint)months.Length);

        var chartWithCustomName = BarChart.Create()
            .WithTitle("Monthly Revenue (Custom Series Name)")
            .WithDataRange(categoriesRange, revenueRange)
            .WithSeriesName("Monthly Revenue")
            .WithPosition(0, 8, 8, 23)
            .WithLegendPosition(ChartLegendPosition.Bottom);

        sheet.AddChart(chartWithCustomName);

        var chartWithDefaultName = BarChart.Create()
            .WithTitle("Monthly Revenue (Default Series Name)")
            .WithDataRange(categoriesRange, revenueRange)
            .WithPosition(9, 8, 17, 23)
            .WithLegendPosition(ChartLegendPosition.Bottom);

        sheet.AddChart(chartWithDefaultName);

        var chartNoLegend = LineChart.Create()
            .WithTitle("Revenue Trend (No Legend)")
            .WithDataRange(categoriesRange, revenueRange)
            .WithPosition(0, 25, 8, 40)
            .WithLegendPosition(ChartLegendPosition.None)
            .WithMarkers(LineChartMarkerStyle.Circle);

        sheet.AddChart(chartNoLegend);

        var multiSeriesChart = BarChart.Create()
            .WithTitle("Revenue vs Expenses (Multiple Series)")
            .WithPosition(9, 25, 17, 40)
            .WithLegendPosition(ChartLegendPosition.Right)
            .AddSeries("Total Revenue", revenueRange)
            .AddSeries("Total Expenses", CellRange.FromBounds(2, 1, 2, (uint)months.Length));

        sheet.AddChart(multiSeriesChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_ChartSeriesNames.xlsx");
    }
}
