using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonMultiColumnAllFeaturesExample : IExample
{
    public string Name => "JSON Multi-Column with All Features";
    public string Description => "Imports JSON array with styling, custom parsers, and auto-fit - showcasing all Phase 1.2 features";


    public int ExampleNumber { get; }

    public JsonMultiColumnAllFeaturesExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Persons.json");
        
        var sheet = WorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Persons Complete")
            .WithHeaderStyle(style => style
                .WithFillColor("70AD47")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithColumnParser("age", value => 
            {
                if (value.IsDecimal() && value.Value.AsT0 > 50)
                    return new($"{value.Value.AsT0} (Senior)");
                return value;
            })
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_JsonMultiColumnAllFeatures.xlsx");
    }
}
