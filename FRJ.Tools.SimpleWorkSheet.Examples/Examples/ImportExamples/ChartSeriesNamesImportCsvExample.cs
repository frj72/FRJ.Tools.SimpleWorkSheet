using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ChartSeriesNamesImportCsvExample : IExample
{
    public string Name => "Chart Series Names Import - CSV";
    public string Description => "Demonstrates series name detection from CSV column headers";

    public void Run()
    {
        var csvData = """
        category,sales,profit
        Product A,150000,45000
        Product B,120000,38000
        Product C,95000,28000
        Product D,180000,54000
        """;

        var workbook = WorkbookBuilder.FromCsv(csvData, hasHeader: true)
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

        ExampleRunner.SaveWorkBook(workbook, "82_ChartSeriesNamesImport_CSV.xlsx");
    }
}
