using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ChartSeriesNamesImportCustomExample : IExample
{
    public string Name => "Chart Series Names Import - Custom Name";
    public string Description => "Demonstrates custom series name override for JSON imports";


    public int ExampleNumber { get; }

    public ChartSeriesNamesImportCustomExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var jsonData = """
        [
            {"month": "Jan", "revenue": 45000, "temperature": 15},
            {"month": "Feb", "revenue": 48000, "temperature": 18},
            {"month": "Mar", "revenue": 52000, "temperature": 22},
            {"month": "Apr", "revenue": 49000, "temperature": 25},
            {"month": "May", "revenue": 54000, "temperature": 28},
            {"month": "Jun", "revenue": 58000, "temperature": 32}
        ]
        """;

        var workbook = WorkbookBuilder.FromJson(jsonData)
            .WithWorkbookName("Custom Series Name")
            .WithDataSheetName("Monthly Data")
            .WithHeaderStyle(style => style
                .WithFillColor("70AD47")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithChart(chart => chart
                .OnSheet("Temperature Chart")
                .UseColumns("month", "temperature")
                .AsLineChart()
                .WithTitle("Monthly Temperature (Custom Series Name)")
                .WithSeriesName("Average Temperature"))
            .WithFreezeHeaderRow()
            .AutoFitAllColumns()
            .Build();

        ExampleRunner.SaveWorkBook(workbook, $"{ExampleNumber:000}_ChartSeriesNamesImport_Custom.xlsx");
    }
}
