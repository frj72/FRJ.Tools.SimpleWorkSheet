using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonFluentImportPersonsExample : IExample
{
    public string Name => "JSON Multi-Column Array Import";
    public string Description => "Fluent import of JSON array with multiple properties per object";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Persons.json");
        
        var sheet = JsonWorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Persons")
            .WithPreserveOriginalValue(false)
            .WithTrimWhitespace(true)
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "56_JsonMultiColumnArrayImport.xlsx");
    }
}
