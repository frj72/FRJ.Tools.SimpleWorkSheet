using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonColumnOrderingExample : IExample
{
    public string Name => "JSON Column Ordering";
    public string Description => "Controls column order in imported JSON data";


    public int ExampleNumber { get; }

    public JsonColumnOrderingExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Persons.json");
        
        var sheet = WorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Persons Ordered")
            .WithColumnOrder("firstName", "lastName", "age", "email")
            .WithHeaderStyle(style => style
                .WithFillColor("5B9BD5")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_JsonColumnOrdering.xlsx");
    }
}
