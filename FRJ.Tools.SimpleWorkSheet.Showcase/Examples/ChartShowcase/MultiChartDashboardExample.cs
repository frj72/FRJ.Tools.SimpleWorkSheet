using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ChartShowcase;

public class MultiChartDashboardExample : IShowcase
{
    public string Name => "Multi-Chart Dashboard";
    public string Description => "4+ different charts on one sheet";
    public string Category => "Chart Showcase";

    public void Run()
    {
        var sheet = new WorkSheet("Dashboard");
        
        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        sheet.AddCell(new(0, 0), "Month", null);
        for (uint i = 0; i < 6; i++)
            sheet.AddCell(new(i + 1, 0), months[i], null);
        
        sheet.AddCell(new(0, 1), "Sales", null);
        var sales = new[] { 100, 120, 115, 130, 125, 140 };
        for (uint i = 0; i < 6; i++)
            sheet.AddCell(new(i + 1, 1), sales[i], null);
        
        sheet.AddCell(new(0, 2), "Costs", null);
        var costs = new[] { 60, 70, 68, 75, 72, 80 };
        for (uint i = 0; i < 6; i++)
            sheet.AddCell(new(i + 1, 2), costs[i], null);
        
        var barChart = BarChart.Create()
            .WithTitle("Monthly Sales")
            .WithDataRange(CellRange.FromBounds(1, 0, 6, 0), CellRange.FromBounds(1, 1, 6, 1))
            .WithPosition(0, 5, 7, 18);
        
        var lineChart = LineChart.Create()
            .WithTitle("Sales Trend")
            .WithDataRange(CellRange.FromBounds(1, 0, 6, 0), CellRange.FromBounds(1, 1, 6, 1))
            .WithPosition(8, 5, 15, 18)
            .WithMarkers(LineChartMarkerStyle.Circle);
        
        var pieChart = PieChart.Create()
            .WithTitle("Revenue Distribution")
            .WithDataRange(CellRange.FromBounds(1, 0, 6, 0), CellRange.FromBounds(1, 1, 6, 1))
            .WithPosition(0, 20, 7, 33);
        
        var scatterChart = ScatterChart.Create()
            .WithTitle("Sales vs Costs")
            .WithXyData(CellRange.FromBounds(1, 1, 6, 1), CellRange.FromBounds(1, 2, 6, 2))
            .WithPosition(8, 20, 15, 33);
        
        sheet.AddChart(barChart);
        sheet.AddChart(lineChart);
        sheet.AddChart(pieChart);
        sheet.AddChart(scatterChart);
        
        var workbook = new WorkBook("MultiChart", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_19_MultiChartDashboard.xlsx");
    }
}
