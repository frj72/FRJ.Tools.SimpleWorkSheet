using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ChartShowcase;

public class ChartFormattingExample : IShowcase
{
    public string Name => "Chart with Custom Formatting";
    public string Description => "Charts with custom colors, labels, and styling";
    public string Category => "Chart Showcase";

    public void Run()
    {
        var sheet = new WorkSheet("FormattedCharts");
        
        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        sheet.AddCell(new(0, 0), "Month", null);
        for (uint i = 0; i < 6; i++)
            sheet.AddCell(new(i + 1, 0), months[i], null);
        
        sheet.AddCell(new(0, 1), "2024", null);
        var data2024 = new[] { 100, 110, 105, 120, 115, 125 };
        for (uint i = 0; i < 6; i++)
            sheet.AddCell(new(i + 1, 1), data2024[i], null);
        
        sheet.AddCell(new(0, 2), "2025", null);
        var data2025 = new[] { 120, 130, 125, 140, 135, 150 };
        for (uint i = 0; i < 6; i++)
            sheet.AddCell(new(i + 1, 2), data2025[i], null);
        
        var chart = BarChart.Create()
            .WithTitle("Year-over-Year Comparison")
            .WithDataRange(CellRange.FromBounds(1, 0, 6, 0), CellRange.FromBounds(1, 1, 6, 1))
            .WithPosition(0, 5, 12, 22)
            .AddSeries("2024", CellRange.FromBounds(1, 1, 6, 1))
            .AddSeries("2025", CellRange.FromBounds(1, 2, 6, 2));
        
        sheet.AddChart(chart);
        
        var workbook = new WorkBook("FormattedChart", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_21_ChartFormatting.xlsx");
    }
}
