using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ChartSeriesNamesImportGenericTableExample : IExample
{
    public string Name => "Chart Series Names Import - Generic Table";
    public string Description => "Demonstrates series name detection from GenericTable column names";

    public void Run()
    {
        var table = GenericTable.Create("Product", "Q1", "Q2", "Q3", "Q4");
        table.AddRow("Widget", 100, 120, 115, 130);
        table.AddRow("Gadget", 80, 95, 105, 110);
        table.AddRow("Tool", 60, 75, 85, 95);

        var workbook = WorkbookBuilder.FromGenericTable(table)
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

        ExampleRunner.SaveWorkBook(workbook, "83_ChartSeriesNamesImport_GenericTable.xlsx");
    }
}
