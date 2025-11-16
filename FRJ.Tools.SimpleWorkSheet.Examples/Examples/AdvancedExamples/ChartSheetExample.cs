using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
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

        dataSheet.AddCell(new(0, 1), new CellValue("Jan"));
        dataSheet.AddCell(new(1, 1), new CellValue(50000));
        dataSheet.AddCell(new(2, 1), new CellValue(30000));
        dataSheet.AddCell(new(3, 1), new CellValue(20000));

        dataSheet.AddCell(new(0, 2), new CellValue("Feb"));
        dataSheet.AddCell(new(1, 2), new CellValue(55000));
        dataSheet.AddCell(new(2, 2), new CellValue(32000));
        dataSheet.AddCell(new(3, 2), new CellValue(23000));

        dataSheet.AddCell(new(0, 3), new CellValue("Mar"));
        dataSheet.AddCell(new(1, 3), new CellValue(60000));
        dataSheet.AddCell(new(2, 3), new CellValue(35000));
        dataSheet.AddCell(new(3, 3), new CellValue(25000));

        dataSheet.AddCell(new(0, 4), new CellValue("Apr"));
        dataSheet.AddCell(new(1, 4), new CellValue(58000));
        dataSheet.AddCell(new(2, 4), new CellValue(33000));
        dataSheet.AddCell(new(3, 4), new CellValue(25000));

        dataSheet.AddCell(new(0, 5), new CellValue("May"));
        dataSheet.AddCell(new(1, 5), new CellValue(62000));
        dataSheet.AddCell(new(2, 5), new CellValue(36000));
        dataSheet.AddCell(new(3, 5), new CellValue(26000));

        dataSheet.AddCell(new(0, 6), new CellValue("Jun"));
        dataSheet.AddCell(new(1, 6), new CellValue(65000));
        dataSheet.AddCell(new(2, 6), new CellValue(38000));
        dataSheet.AddCell(new(3, 6), new CellValue(27000));

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
            "45_ChartSheet.xlsx");

        workbook.SaveToFile(outputPath);

        Console.WriteLine($"Saved: {outputPath}");
        Console.WriteLine("✓ Completed successfully");
    }
}
