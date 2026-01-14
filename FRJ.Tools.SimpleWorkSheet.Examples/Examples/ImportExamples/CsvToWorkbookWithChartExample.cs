using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class CsvToWorkbookWithChartExample : IExample
{
    public string Name => "CSV to Workbook with Chart";
    public string Description => "Creates workbook with data sheet and chart from CSV file";

    public void Run()
    {
        var csvPath = Path.Combine("Resources", "Data", "Csv", "MixedTypes.csv");
        
        var workbook = WorkbookBuilder.FromCsvFile(csvPath)
            .WithWorkbookName("CSV Analysis Report")
            .WithDataSheetName("Raw Data")
            .WithHeaderStyle(style => style
                .WithFillColor("5B9BD5")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithFreezeHeaderRow()
            .WithChart(chart => chart
                .OnSheet("Amount Trend")
                .UseColumns("signupDate", "amount")
                .AsLineChart()
                .WithTitle("Amount Over Time")
                .WithCategoryAxisTitle("Signup Date")
                .WithValueAxisTitle("Amount")
                .WithLegendPosition(ChartLegendPosition.Bottom))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, "073_CsvToWorkbookWithChart.xlsx");
    }
}
