using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonColumnOrderingExample : IExample
{
    public string Name => "JSON Column Ordering";
    public string Description => "Controls column order in imported JSON data";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Persons.json");
        
        var sheet = JsonWorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Persons Ordered")
            .WithColumnOrder("firstName", "lastName", "age", "email")
            .WithHeaderStyle(style => style
                .WithFillColor("5B9BD5")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "65_JsonColumnOrdering.xlsx");
    }
}
