using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class LineChartWithYAxisLabelsExample : IExample
{
    public string Name => "Line Chart With Explicit Y-Axis Labels";
    public string Description => "Demonstrates line charts with y-axis tick labels explicitly enabled";

    public void Run()
    {
        var sheet = new WorkSheet("Revenue Trend");

        sheet.AddCell(new(0, 0), "Month", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Revenue", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug" };
        var revenue = new[] { 45000, 52000, 48000, 58000, 62000, 59000, 67000, 71000 };

        for (uint i = 0; i < months.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), months[i], null);
            sheet.AddCell(new(1, i + 1), revenue[i], null);
        }

        var categoriesRange = CellRange.FromBounds(0, 1, 0, (uint)months.Length);
        var valuesRange = CellRange.FromBounds(1, 1, 1, (uint)months.Length);

        var lineChart = LineChart.Create()
            .WithTitle("Monthly Revenue with Y-Axis Labels")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(0, 10, 10, 27)
            .WithMarkers(LineChartMarkerStyle.Circle)
            .WithCategoryAxisTitle("Month")
            .WithValueAxisTitle("Revenue ($)")
            .WithYAxisLabels(true);

        sheet.AddChart(lineChart);

        ExampleRunner.SaveWorkSheet(sheet, "097_LineChartWithExplicitYAxisLabels.xlsx");
    }
}
