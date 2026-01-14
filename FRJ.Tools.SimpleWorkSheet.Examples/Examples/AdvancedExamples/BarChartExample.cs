using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class BarChartExample : IExample
{
    public string Name => "Bar Chart";
    public string Description => "Demonstrates creating bar charts with data visualization";

    public void Run()
    {
        var sheet = new WorkSheet("Sales Data");

        sheet.AddCell(new(0, 0), "Region", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Q1 Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "Q2 Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(3, 0), "Q3 Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(4, 0), "Q4 Sales", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.AddCell(new(0, 1), new("North"), null);
        sheet.AddCell(new(1, 1), new(125000), null);
        sheet.AddCell(new(2, 1), new(138000), null);
        sheet.AddCell(new(3, 1), new(142000), null);
        sheet.AddCell(new(4, 1), new(155000), null);

        sheet.AddCell(new(0, 2), new("South"), null);
        sheet.AddCell(new(1, 2), new(98000), null);
        sheet.AddCell(new(2, 2), new(105000), null);
        sheet.AddCell(new(3, 2), new(112000), null);
        sheet.AddCell(new(4, 2), new(118000), null);

        sheet.AddCell(new(0, 3), new("East"), null);
        sheet.AddCell(new(1, 3), new(87000), null);
        sheet.AddCell(new(2, 3), new(92000), null);
        sheet.AddCell(new(3, 3), new(96000), null);
        sheet.AddCell(new(4, 3), new(101000), null);

        sheet.AddCell(new(0, 4), new("West"), null);
        sheet.AddCell(new(1, 4), new(145000), null);
        sheet.AddCell(new(2, 4), new(152000), null);
        sheet.AddCell(new(3, 4), new(158000), null);
        sheet.AddCell(new(4, 4), new(165000), null);

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 4);
        var q1Range = CellRange.FromBounds(1, 1, 1, 4);

        var verticalChart = BarChart.Create()
            .WithTitle("Q1 Sales by Region")
            .WithDataRange(categoriesRange, q1Range)
            .WithPosition(0, 6, 8, 21)
            .WithOrientation(BarChartOrientation.Vertical);

        sheet.AddChart(verticalChart);

        var q4Range = CellRange.FromBounds(4, 1, 4, 4);

        var horizontalChart = BarChart.Create()
            .WithTitle("Q4 Sales by Region (Horizontal)")
            .WithDataRange(categoriesRange, q4Range)
            .WithPosition(9, 6, 17, 21)
            .WithOrientation(BarChartOrientation.Horizontal);

        sheet.AddChart(horizontalChart);

        ExampleRunner.SaveWorkSheet(sheet, "041_BarChart.xlsx");
    }
}
