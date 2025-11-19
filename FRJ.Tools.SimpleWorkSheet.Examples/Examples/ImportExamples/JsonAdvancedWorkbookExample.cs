using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonAdvancedWorkbookExample : IExample
{
    public string Name => "JSON Advanced Workbook Features";
    public string Description => "Showcases frozen headers, custom chart positioning, sizing, and data labels";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices.json");
        
        var workbook = WorkbookBuilder.FromJsonFile(jsonPath)
            .WithWorkbookName("Advanced Price Analysis")
            .WithDataSheetName("Price Data")
            .WithHeaderStyle(style => style
                .WithFillColor("4472C4")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithFreezeHeaderRow()
            .WithChart(chart => chart
                .OnSheet("Price Trend")
                .UseColumns("Date", "Price")
                .AsLineChart()
                .WithTitle("Price Over Time")
                .WithCategoryAxisTitle("Date")
                .WithValueAxisTitle("Price")
                .WithChartPosition(0, 0, 10, 20)
                .WithChartSize(600, 400)
                .WithDataLabels(false)
                .WithLegendPosition(ChartLegendPosition.Bottom))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, "69_JsonAdvancedWorkbook.xlsx");
    }
}
