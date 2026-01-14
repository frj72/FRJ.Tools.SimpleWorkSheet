using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ChartSeriesNamesImportAutoExample : IExample
{
    public string Name => "Chart Series Names Import - Auto Detection";
    public string Description => "Demonstrates automatic series name detection from JSON column names";

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
            .WithWorkbookName("Auto Series Name")
            .WithDataSheetName("Monthly Data")
            .WithHeaderStyle(style => style
                .WithFillColor("4472C4")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithChart(chart => chart
                .OnSheet("Revenue Chart")
                .UseColumns("month", "revenue")
                .AsBarChart()
                .WithTitle("Monthly Revenue (Auto-Detected Series Name)"))
            .WithFreezeHeaderRow()
            .AutoFitAllColumns()
            .Build();

        ExampleRunner.SaveWorkBook(workbook, "080_ChartSeriesNamesImport_Auto.xlsx");
    }
}
