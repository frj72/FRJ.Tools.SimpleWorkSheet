using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class LineChartMultipleSeriesWithYAxisLabelsExample : IExample
{
    public string Name => "Line Chart Multiple Series With Explicit Y-Axis Labels";
    public string Description => "Demonstrates multi-series line charts with y-axis tick labels explicitly enabled";


    public int ExampleNumber { get; }

    public LineChartMultipleSeriesWithYAxisLabelsExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Department Performance");

        sheet.AddCell(new(0, 0), "Quarter", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Engineering", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(3, 0), "Marketing", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var quarters = new[] { "Q1", "Q2", "Q3", "Q4", "Q5", "Q6" };
        var engineering = new[] { 185000, 192000, 198000, 205000, 215000, 225000 };
        var sales = new[] { 145000, 158000, 172000, 185000, 195000, 208000 };
        var marketing = new[] { 95000, 102000, 108000, 115000, 122000, 130000 };

        for (uint i = 0; i < quarters.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), quarters[i], null);
            sheet.AddCell(new(1, i + 1), engineering[i], null);
            sheet.AddCell(new(2, i + 1), sales[i], null);
            sheet.AddCell(new(3, i + 1), marketing[i], null);
        }

        var engineeringRange = CellRange.FromBounds(1, 1, 1, (uint)quarters.Length);
        var salesRange = CellRange.FromBounds(2, 1, 2, (uint)quarters.Length);
        var marketingRange = CellRange.FromBounds(3, 1, 3, (uint)quarters.Length);

        var lineChart = LineChart.Create()
            .WithTitle("Department Performance with Y-Axis Labels")
            .WithPosition(0, 8, 12, 25)
            .WithMarkers(LineChartMarkerStyle.Circle)
            .WithCategoryAxisTitle("Quarter")
            .WithValueAxisTitle("Budget ($)")
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .WithYAxisLabels(true)
            .AddSeries("Engineering", engineeringRange)
            .AddSeries("Sales", salesRange)
            .AddSeries("Marketing", marketingRange);

        sheet.AddChart(lineChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_LineChartMultipleSeriesWithExplicitYAxisLabels.xlsx");
    }
}
