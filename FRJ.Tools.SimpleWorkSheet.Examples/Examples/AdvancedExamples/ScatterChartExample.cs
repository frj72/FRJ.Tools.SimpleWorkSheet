using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ScatterChartExample : IExample
{
    public string Name => "Scatter Chart";
    public string Description => "Demonstrates scatter charts for correlation analysis";

    public void Run()
    {
        var sheet = new WorkSheet("Correlation Data");

        sheet.AddCell(new(0, 0), "Study Hours", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Test Score", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.AddCell(new(0, 1), new(2), null);
        sheet.AddCell(new(1, 1), new(55), null);

        sheet.AddCell(new(0, 2), new(3), null);
        sheet.AddCell(new(1, 2), new(62), null);

        sheet.AddCell(new(0, 3), new(4), null);
        sheet.AddCell(new(1, 3), new(68), null);

        sheet.AddCell(new(0, 4), new(5), null);
        sheet.AddCell(new(1, 4), new(75), null);

        sheet.AddCell(new(0, 5), new(6), null);
        sheet.AddCell(new(1, 5), new(80), null);

        sheet.AddCell(new(0, 6), new(7), null);
        sheet.AddCell(new(1, 6), new(85), null);

        sheet.AddCell(new(0, 7), new(8), null);
        sheet.AddCell(new(1, 7), new(90), null);

        var xRange = CellRange.FromBounds(0, 1, 0, 7);
        var yRange = CellRange.FromBounds(1, 1, 1, 7);

        var scatterChart = ScatterChart.Create()
            .WithTitle("Study Hours vs Test Score")
            .WithXyData(xRange, yRange)
            .WithPosition(0, 9, 8, 24)
            .WithTrendline(true);

        sheet.AddChart(scatterChart);

        ExampleRunner.SaveWorkSheet(sheet, "044_ScatterChart.xlsx");
    }
}
