using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ChartSeriesNamesImportExample : IExample
{
    public string Name => "Chart Series Names Import";
    public string Description => "Demonstrates automatic column name detection in chart legends for imported data";

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

        var workbook1 = WorkbookBuilder.FromJson(jsonData)
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

        ExampleRunner.SaveWorkBook(workbook1, "80_ChartSeriesNamesImport_Auto.xlsx");

        var workbook2 = WorkbookBuilder.FromJson(jsonData)
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

        ExampleRunner.SaveWorkBook(workbook2, "81_ChartSeriesNamesImport_Custom.xlsx");

        var csvData = """
        category,sales,profit
        Product A,150000,45000
        Product B,120000,38000
        Product C,95000,28000
        Product D,180000,54000
        """;

        var workbook3 = WorkbookBuilder.FromCsv(csvData, hasHeader: true)
            .WithWorkbookName("CSV Chart Import")
            .WithDataSheetName("Sales Data")
            .WithHeaderStyle(style => style
                .WithFillColor("ED7D31")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithChart(chart => chart
                .OnSheet("Sales Chart")
                .UseColumns("category", "sales")
                .AsPieChart()
                .WithTitle("Sales Distribution (Column: sales)")
                .WithDataLabels(true))
            .AutoFitAllColumns()
            .Build();

        ExampleRunner.SaveWorkBook(workbook3, "82_ChartSeriesNamesImport_CSV.xlsx");

        var table = GenericTable.Create("Product", "Q1", "Q2", "Q3", "Q4");
        table.AddRow("Widget", 100, 120, 115, 130);
        table.AddRow("Gadget", 80, 95, 105, 110);
        table.AddRow("Tool", 60, 75, 85, 95);

        var workbook4 = WorkbookBuilder.FromGenericTable(table)
            .WithWorkbookName("Generic Table Chart")
            .WithDataSheetName("Quarterly Sales")
            .WithHeaderStyle(style => style
                .WithFillColor("5B9BD5")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithChart(chart => chart
                .OnSheet("Q1 Sales")
                .UseColumns("Product", "Q1")
                .AsBarChart()
                .WithTitle("Q1 Sales by Product (Column: Q1)")
                .WithLegendPosition(ChartLegendPosition.Bottom))
            .AutoFitAllColumns()
            .Build();

        ExampleRunner.SaveWorkBook(workbook4, "83_ChartSeriesNamesImport_GenericTable.xlsx");
    }
}
