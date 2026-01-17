using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class LineChartMultipleSeriesWithoutCategoriesExample : IExample
{
    public string Name => "Line Chart Multiple Series Without Categories";
    public string Description => "Demonstrates line charts with multiple series without explicit x-axis categories";


    public int ExampleNumber { get; }

    public LineChartMultipleSeriesWithoutCategoriesExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Metrics Data");

        var metric1 = new[] { 45, 52, 48, 55, 58, 62, 59, 64 };
        var metric2 = new[] { 38, 42, 44, 47, 50, 53, 51, 56 };
        var metric3 = new[] { 41, 43, 46, 48, 51, 55, 54, 58 };

        for (uint i = 0; i < metric1.Length; i++)
        {
            sheet.AddCell(new(0, i), new(metric1[i]), null);
            sheet.AddCell(new(1, i), new(metric2[i]), null);
            sheet.AddCell(new(2, i), new(metric3[i]), null);
        }

        var series1Range = CellRange.FromBounds(0, 0, 0, (uint)metric1.Length - 1);
        var series2Range = CellRange.FromBounds(1, 0, 1, (uint)metric2.Length - 1);
        var series3Range = CellRange.FromBounds(2, 0, 2, (uint)metric3.Length - 1);

        var lineChart = LineChart.Create()
            .WithTitle("Performance Metrics Comparison")
            .WithPosition(4, 0, 14, 16)
            .WithMarkers(LineChartMarkerStyle.Circle)
            .WithValueAxisTitle("Value")
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .AddSeries("Metric A", series1Range)
            .AddSeries("Metric B", series2Range)
            .AddSeries("Metric C", series3Range);

        sheet.AddChart(lineChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_LineChartMultipleSeriesWithoutCategories.xlsx");
    }
}
