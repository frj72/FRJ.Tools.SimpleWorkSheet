using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class LineChartMultipleSeriesWithDataLabelsExample : IExample
{
    public string Name => "Line Chart Multiple Series With Data Labels";
    public string Description => "Demonstrates line charts with multiple series and data labels";

    public void Run()
    {
        var sheet = new WorkSheet("Regional Comparison");

        sheet.AddCell(new(0, 0), "Quarter", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "North", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "South", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(3, 0), "West", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var quarters = new[] { "Q1", "Q2", "Q3", "Q4", "Q5", "Q6" };
        var northSales = new[] { 125000, 142000, 138000, 155000, 168000, 175000 };
        var southSales = new[] { 98000, 115000, 122000, 135000, 142000, 158000 };
        var westSales = new[] { 110000, 128000, 135000, 148000, 156000, 165000 };

        for (uint i = 0; i < quarters.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), quarters[i], null);
            sheet.AddCell(new(1, i + 1), northSales[i], null);
            sheet.AddCell(new(2, i + 1), southSales[i], null);
            sheet.AddCell(new(3, i + 1), westSales[i], null);
        }

        var northRange = CellRange.FromBounds(1, 1, 1, (uint)quarters.Length);
        var southRange = CellRange.FromBounds(2, 1, 2, (uint)quarters.Length);
        var westRange = CellRange.FromBounds(3, 1, 3, (uint)quarters.Length);

        var lineChart = LineChart.Create()
            .WithTitle("Regional Sales Comparison with Data Labels")
            .WithPosition(0, 8, 12, 25)
            .WithMarkers(LineChartMarkerStyle.Circle)
            .WithCategoryAxisTitle("Quarter")
            .WithValueAxisTitle("Sales ($)")
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .WithDataLabels(true)
            .AddSeries("North Region", northRange)
            .AddSeries("South Region", southRange)
            .AddSeries("West Region", westRange);

        sheet.AddChart(lineChart);

        ExampleRunner.SaveWorkSheet(sheet, "95_LineChartMultipleSeriesWithDataLabels.xlsx");
    }
}
