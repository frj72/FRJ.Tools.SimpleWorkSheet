using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonToWorkbookLineChartExample : IExample
{
    public string Name => "JSON to Workbook with Line Chart";
    public string Description => "Creates workbook with data sheet and line chart using column names from JSON";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices.json");
        
        var workbook = WorkbookBuilder.FromJsonFile(jsonPath)
            .WithWorkbookName("Price Trend Analysis")
            .WithDataSheetName("Historical Prices")
            .WithHeaderStyle(style => style
                .WithFillColor("4472C4")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithChart(chart => chart
                .OnSheet("Price Trend")
                .UseColumns("Date", "Price")
                .AsLineChart()
                .WithTitle("Price Over Time"))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, "62_JsonToWorkbookLineChart.xlsx");
    }
}
