using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class LineChartExample : IExample
{
    public string Name => "Line Chart";
    public string Description => "Demonstrates line charts for trend visualization";

    public void Run()
    {
        var sheet = new WorkSheet("Monthly Sales");

        sheet.AddCell(new(0, 0), "Month", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "2023 Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "2024 Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.AddCell(new(0, 1), new("Jan"), null);
        sheet.AddCell(new(1, 1), new(45000), null);
        sheet.AddCell(new(2, 1), new(52000), null);

        sheet.AddCell(new(0, 2), new("Feb"), null);
        sheet.AddCell(new(1, 2), new(48000), null);
        sheet.AddCell(new(2, 2), new(55000), null);

        sheet.AddCell(new(0, 3), new("Mar"), null);
        sheet.AddCell(new(1, 3), new(52000), null);
        sheet.AddCell(new(2, 3), new(58000), null);

        sheet.AddCell(new(0, 4), new("Apr"), null);
        sheet.AddCell(new(1, 4), new(49000), null);
        sheet.AddCell(new(2, 4), new(61000), null);

        sheet.AddCell(new(0, 5), new("May"), null);
        sheet.AddCell(new(1, 5), new(54000), null);
        sheet.AddCell(new(2, 5), new(64000), null);

        sheet.AddCell(new(0, 6), new("Jun"), null);
        sheet.AddCell(new(1, 6), new(58000), null);
        sheet.AddCell(new(2, 6), new(67000), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 6);
        var values2024Range = CellRange.FromBounds(2, 1, 2, 6);

        var lineChart = LineChart.Create()
            .WithTitle("Monthly Sales Trend (2024)")
            .WithDataRange(categoriesRange, values2024Range)
            .WithPosition(0, 8, 8, 23)
            .WithMarkers(LineChartMarkerStyle.Circle);

        sheet.AddChart(lineChart);

        ExampleRunner.SaveWorkSheet(sheet, "42_LineChart.xlsx");
    }
}
