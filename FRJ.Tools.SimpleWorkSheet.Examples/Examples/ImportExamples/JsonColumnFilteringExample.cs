using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonColumnFilteringExample : IExample
{
    public string Name => "JSON Column Filtering";
    public string Description => "Includes/excludes specific columns from imported JSON data";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Orders.json");
        
        var sheet = JsonWorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Orders Filtered")
            .WithIncludeColumns("orderId", "orderDate", "totalAmount", "status")
            .WithHeaderStyle(style => style
                .WithFillColor("ED7D31")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "66_JsonColumnFiltering.xlsx");
    }
}
