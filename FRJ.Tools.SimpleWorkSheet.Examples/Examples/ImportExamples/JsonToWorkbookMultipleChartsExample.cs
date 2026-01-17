using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonToWorkbookMultipleChartsExample : IExample
{
    public string Name => "JSON to Workbook with Multiple Charts";
    public string Description => "Creates workbook with data sheet and multiple charts (pie + area) with full formatting";


    public int ExampleNumber { get; }

    public JsonToWorkbookMultipleChartsExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Orders.json");
        
        var workbook = WorkbookBuilder.FromJsonFile(jsonPath)
            .WithWorkbookName("Order Analysis")
            .WithDataSheetName("Orders")
            .WithHeaderStyle(style => style
                .WithFillColor("ED7D31")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithChart(chart => chart
                .OnSheet("Status Distribution")
                .UseColumns("status", "totalAmount")
                .AsPieChart()
                .WithTitle("Orders by Status")
                .WithLegendPosition(ChartLegendPosition.Right))
            .WithChart(chart => chart
                .OnSheet("Revenue Trend")
                .UseColumns("orderDate", "totalAmount")
                .AsAreaChart()
                .WithTitle("Revenue Over Time")
                .WithCategoryAxisTitle("Date")
                .WithValueAxisTitle("Amount ($)")
                .WithLegendPosition(ChartLegendPosition.Bottom))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, $"{ExampleNumber:000}_JsonToWorkbookMultipleCharts.xlsx");
    }
}
