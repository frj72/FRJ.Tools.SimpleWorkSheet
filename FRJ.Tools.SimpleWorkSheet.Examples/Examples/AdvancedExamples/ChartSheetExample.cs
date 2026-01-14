using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ChartSheetExample : IExample
{
    public string Name => "Chart Sheet";
    public string Description => "Demonstrates creating a dedicated chart dashboard sheet that references data from another sheet";

    public void Run()
    {
        var dataSheet = new WorkSheet("Data");

        dataSheet.AddCell(new(0, 0), "Month", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        dataSheet.AddCell(new(1, 0), "Revenue", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        dataSheet.AddCell(new(2, 0), "Expenses", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        dataSheet.AddCell(new(3, 0), "Profit", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        dataSheet.AddCell(new(0, 1), new("Jan"), null);
        dataSheet.AddCell(new(1, 1), new(50000), null);
        dataSheet.AddCell(new(2, 1), new(30000), null);
        dataSheet.AddCell(new(3, 1), new(20000), null);

        dataSheet.AddCell(new(0, 2), new("Feb"), null);
        dataSheet.AddCell(new(1, 2), new(55000), null);
        dataSheet.AddCell(new(2, 2), new(32000), null);
        dataSheet.AddCell(new(3, 2), new(23000), null);

        dataSheet.AddCell(new(0, 3), new("Mar"), null);
        dataSheet.AddCell(new(1, 3), new(60000), null);
        dataSheet.AddCell(new(2, 3), new(35000), null);
        dataSheet.AddCell(new(3, 3), new(25000), null);

        dataSheet.AddCell(new(0, 4), new("Apr"), null);
        dataSheet.AddCell(new(1, 4), new(58000), null);
        dataSheet.AddCell(new(2, 4), new(33000), null);
        dataSheet.AddCell(new(3, 4), new(25000), null);

        dataSheet.AddCell(new(0, 5), new("May"), null);
        dataSheet.AddCell(new(1, 5), new(62000), null);
        dataSheet.AddCell(new(2, 5), new(36000), null);
        dataSheet.AddCell(new(3, 5), new(26000), null);

        dataSheet.AddCell(new(0, 6), new("Jun"), null);
        dataSheet.AddCell(new(1, 6), new(65000), null);
        dataSheet.AddCell(new(2, 6), new(38000), null);
        dataSheet.AddCell(new(3, 6), new(27000), null);

        var dashboardSheet = new WorkSheet("Dashboard");

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 6);
        var revenueRange = CellRange.FromBounds(1, 1, 1, 6);
        var expensesRange = CellRange.FromBounds(2, 1, 2, 6);
        var profitRange = CellRange.FromBounds(3, 1, 3, 6);

        var revenueChart = BarChart.Create()
            .WithTitle("Monthly Revenue")
            .WithDataRange(categoriesRange, revenueRange)
            .WithPosition(0, 0, 8, 15)
            .WithOrientation(BarChartOrientation.Vertical)
            .WithDataSourceSheet("Data");

        dashboardSheet.AddChart(revenueChart);

        var profitTrendChart = LineChart.Create()
            .WithTitle("Profit Trend")
            .WithDataRange(categoriesRange, profitRange)
            .WithPosition(9, 0, 17, 15)
            .WithMarkers(LineChartMarkerStyle.Circle)
            .WithDataSourceSheet("Data");

        dashboardSheet.AddChart(profitTrendChart);

        var expensesChart = BarChart.Create()
            .WithTitle("Monthly Expenses")
            .WithDataRange(categoriesRange, expensesRange)
            .WithPosition(0, 16, 8, 31)
            .WithOrientation(BarChartOrientation.Horizontal)
            .WithDataSourceSheet("Data");

        dashboardSheet.AddChart(expensesChart);

        var workbook = new WorkBook("Financial Dashboard", [dataSheet, dashboardSheet]);

        var outputPath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "Output",
            "045_ChartSheet.xlsx");

        workbook.SaveToFile(outputPath);

        Console.WriteLine($"Saved: {outputPath}");
        Console.WriteLine("✓ Completed successfully");
    }
}
