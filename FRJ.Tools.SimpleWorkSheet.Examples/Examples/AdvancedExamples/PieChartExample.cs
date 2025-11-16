using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class PieChartExample : IExample
{
    public string Name => "Pie Chart";
    public string Description => "Demonstrates pie charts for distribution visualization";

    public void Run()
    {
        var sheet = new WorkSheet("Market Share");

        sheet.AddCell(new(0, 0), "Company", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Market Share %", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.AddCell(new(0, 1), new CellValue("Company A"));
        sheet.AddCell(new(1, 1), new CellValue(35));

        sheet.AddCell(new(0, 2), new CellValue("Company B"));
        sheet.AddCell(new(1, 2), new CellValue(25));

        sheet.AddCell(new(0, 3), new CellValue("Company C"));
        sheet.AddCell(new(1, 3), new CellValue(20));

        sheet.AddCell(new(0, 4), new CellValue("Company D"));
        sheet.AddCell(new(1, 4), new CellValue(12));

        sheet.AddCell(new(0, 5), new CellValue("Others"));
        sheet.AddCell(new(1, 5), new CellValue(8));

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 5);

        var pieChart = PieChart.Create()
            .WithTitle("Market Share Distribution")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(0, 7, 8, 22)
            .WithExplosion(10);

        sheet.AddChart(pieChart);

        ExampleRunner.SaveWorkSheet(sheet, "43_PieChart.xlsx");
    }
}
