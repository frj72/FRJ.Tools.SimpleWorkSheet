using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonToWorkbookBarChartExample : IExample
{
    public string Name => "JSON to Workbook with Bar Chart";
    public string Description => "Creates workbook with data sheet and bar chart using column names from JSON";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Persons.json");
        
        var workbook = WorkbookBuilder.FromJsonFile(jsonPath)
            .WithWorkbookName("Person Data Analysis")
            .WithDataSheetName("Persons")
            .WithHeaderStyle(style => style
                .WithFillColor("70AD47")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithChart(chart => chart
                .OnSheet("Age Distribution")
                .UseColumns("firstName", "age")
                .AsBarChart()
                .WithTitle("Age by Person"))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, "63_JsonToWorkbookBarChart.xlsx");
    }
}
