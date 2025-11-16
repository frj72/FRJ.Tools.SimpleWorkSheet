using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ChartShowcase;

public class AllChartTypesExample : IShowcase
{
    public string Name => "All Chart Types on Same Data";
    public string Description => "Bar, Line, Pie, Scatter charts using identical data";
    public string Category => "Chart Showcase";

    public void Run()
    {
        var dataSheet = new WorkSheet("Data");
        var barSheet = new WorkSheet("Bar Chart");
        var lineSheet = new WorkSheet("Line Chart");
        var pieSheet = new WorkSheet("Pie Chart");
        var scatterSheet = new WorkSheet("Scatter Chart");
        
        var categories = new[] { "ProductA", "Product B", "Product C", "Product D", "Product E" };
        var values = new[] { 120, 95, 140, 85, 110 };
        
        dataSheet.AddCell(new(0, 0), "Product");
        dataSheet.AddCell(new(1, 0), "Sales");
        for (uint i = 0; i < 5; i++)
        {
            dataSheet.AddCell(new(0, i + 1), categories[i]);
            dataSheet.AddCell(new(1, i + 1), values[i]);
        }
        
        var dataRange = CellRange.FromBounds(0, 1, 0, 5);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 5);
        
        var barChart = BarChart.Create()
            .WithTitle("Sales by Product - Bar Chart")
            .WithDataRange(dataRange, valuesRange)
            .WithDataSourceSheet("Data")
            .WithPosition(3, 1, 15, 20);
        barSheet.AddChart(barChart);
        
        var lineChart = LineChart.Create()
            .WithTitle("Sales by Product - Line Chart")
            .WithDataRange(dataRange, valuesRange)
            .WithDataSourceSheet("Data")
            .WithPosition(3, 1, 15, 20)
            .WithMarkers(LineChartMarkerStyle.Diamond);
        lineSheet.AddChart(lineChart);
        
        var pieChart = PieChart.Create()
            .WithTitle("Sales by Product - Pie Chart")
            .WithDataRange(dataRange, valuesRange)
            .WithDataSourceSheet("Data")
            .WithPosition(3, 1, 15, 20);
        pieSheet.AddChart(pieChart);
        
        var scatterChart = ScatterChart.Create()
            .WithTitle("Sales Distribution - Scatter")
            .WithXyData(valuesRange, valuesRange)
            .WithDataSourceSheet("Data")
            .WithPosition(3, 1, 15, 20);
        scatterSheet.AddChart(scatterChart);
        
        var workbook = new WorkBook("AllCharts", [dataSheet, barSheet, lineSheet, pieSheet, scatterSheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_20_AllChartTypes.xlsx");
    }
}
