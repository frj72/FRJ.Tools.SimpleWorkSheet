using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.StressTests;

public class LargeChartExample : IShowcase
{
    public string Name => "Chart with 1000+ Data Points";
    public string Description => "Tests chart performance with large datasets";
    public string Category => "Stress Tests";

    public void Run()
    {
        var sheet = new WorkSheet("ChartData");
        
        sheet.AddCell(new(0, 0), "X", null);
        sheet.AddCell(new(1, 0), "Y", null);
        
        Console.WriteLine("  Creating 1000 data points...");
        for (uint row = 1; row <= 1000; row++)
        {
            sheet.AddCell(new(0, row), row, null);
            sheet.AddCell(new(1, row), Math.Sin(row / 10.0) * 100, null);
            
            if (row % 100 == 0)
                Console.WriteLine($"    Created {row} points...");
        }
        
        var chart = LineChart.Create()
            .WithTitle("Sine Wave - 1000 Points")
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 1000),
                CellRange.FromBounds(1, 1, 1, 1000))
            .WithPosition(3, 1, 15, 20)
            .WithMarkers(LineChartMarkerStyle.None)
            .WithSmoothLines(true);
        
        sheet.AddChart(chart);
        
        var workbook = new WorkBook("LargeChart", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_12_LargeChart.xlsx");
    }
}
