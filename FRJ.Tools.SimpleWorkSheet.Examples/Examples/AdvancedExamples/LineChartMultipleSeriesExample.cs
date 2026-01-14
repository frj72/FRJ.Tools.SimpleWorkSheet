using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class LineChartMultipleSeriesExample : IExample
{
    public string Name => "Line Chart Multiple Series";
    public string Description => "Demonstrates line charts with multiple data series";

    public void Run()
    {
        var sheet = new WorkSheet("Regional Sales");

        sheet.AddCell(new(0, 0), "Month", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "North Region", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "South Region", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(3, 0), "West Region", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        var northSales = new[] { 45000, 48000, 52000, 49000, 54000, 58000 };
        var southSales = new[] { 38000, 42000, 44000, 47000, 50000, 53000 };
        var westSales = new[] { 41000, 43000, 46000, 48000, 51000, 55000 };

        for (uint i = 0; i < months.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), months[i], null);
            sheet.AddCell(new(1, i + 1), northSales[i], null);
            sheet.AddCell(new(2, i + 1), southSales[i], null);
            sheet.AddCell(new(3, i + 1), westSales[i], null);
        }

        var northRange = CellRange.FromBounds(1, 1, 1, (uint)months.Length);
        var southRange = CellRange.FromBounds(2, 1, 2, (uint)months.Length);
        var westRange = CellRange.FromBounds(3, 1, 3, (uint)months.Length);

        var lineChart = LineChart.Create()
            .WithTitle("Regional Sales Comparison")
            .WithPosition(0, 8, 12, 25)
            .WithMarkers(LineChartMarkerStyle.Circle)
            .WithCategoryAxisTitle("Month")
            .WithValueAxisTitle("Sales ($)")
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .AddSeries("North Region", northRange)
            .AddSeries("South Region", southRange)
            .AddSeries("West Region", westRange);

        sheet.AddChart(lineChart);

        ExampleRunner.SaveWorkSheet(sheet, "089_LineChartMultipleSeries.xlsx");
    }
}
