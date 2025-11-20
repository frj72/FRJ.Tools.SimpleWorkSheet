using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonNestedObjectsExample : IExample
{
    public string Name => "JSON Nested Objects Import";
    public string Description => "Imports JSON with nested objects, flattening them with dot notation (e.g., address.city)";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Orders.json");
        
        var sheet = WorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Orders")
            .WithHeaderStyle(style => style
                .WithFillColor("ED7D31")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "60_JsonNestedObjects.xlsx");
    }
}
