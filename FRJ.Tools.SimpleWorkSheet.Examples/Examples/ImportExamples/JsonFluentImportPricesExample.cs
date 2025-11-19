using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonFluentImportPricesExample : IExample
{
    public string Name => "JSON Array Fluent Import";
    public string Description => "Simple fluent import of JSON array with automatic schema discovery and type detection";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices.json");
        
        var sheet = JsonWorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Price Data")
            .WithPreserveOriginalValue(true)
            .WithTrimWhitespace(true)
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "54_JsonArrayFluentImport.xlsx");
    }
}
