using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
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

        sheet.AddCell(new(0, 1), new CellValue(2));
        sheet.AddCell(new(1, 1), new CellValue(55));

        sheet.AddCell(new(0, 2), new CellValue(3));
        sheet.AddCell(new(1, 2), new CellValue(62));

        sheet.AddCell(new(0, 3), new CellValue(4));
        sheet.AddCell(new(1, 3), new CellValue(68));

        sheet.AddCell(new(0, 4), new CellValue(5));
        sheet.AddCell(new(1, 4), new CellValue(75));

        sheet.AddCell(new(0, 5), new CellValue(6));
        sheet.AddCell(new(1, 5), new CellValue(80));

        sheet.AddCell(new(0, 6), new CellValue(7));
        sheet.AddCell(new(1, 6), new CellValue(85));

        sheet.AddCell(new(0, 7), new CellValue(8));
        sheet.AddCell(new(1, 7), new CellValue(90));

        var xRange = CellRange.FromBounds(0, 1, 0, 7);
        var yRange = CellRange.FromBounds(1, 1, 1, 7);

        var scatterChart = ScatterChart.Create()
            .WithTitle("Study Hours vs Test Score")
            .WithXyData(xRange, yRange)
            .WithPosition(0, 9, 8, 24)
            .WithTrendline(true);

        sheet.AddChart(scatterChart);

        ExampleRunner.SaveWorkSheet(sheet, "44_ScatterChart.xlsx");
    }
}
