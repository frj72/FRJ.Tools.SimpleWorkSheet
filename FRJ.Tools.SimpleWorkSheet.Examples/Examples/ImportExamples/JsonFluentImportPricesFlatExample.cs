using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonFluentImportPricesFlatExample : IExample
{
    public string Name => "JSON Flat Object Import";
    public string Description => "Import flat JSON object (key-value pairs) as single row with property names as headers";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices_flat.json");
        
        var sheet = JsonWorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Price Data Flat")
            .WithPreserveOriginalValue(false)
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "55_JsonFlatObjectImport.xlsx");
    }
}
