using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonFlatObjectWithParsersExample : IExample
{
    public string Name => "JSON Flat Object with Column Parsers";
    public string Description => "Imports flat JSON object with custom column parsers to format values";


    public int ExampleNumber { get; }

    public JsonFlatObjectWithParsersExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices_flat.json");
        
        var sheet = WorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Prices Parsed")
            .WithColumnParser("price_1", value => new(value.Value.AsT0 * 100))
            .WithColumnParser("price_10", value => new(value.Value.AsT0 * 100))
            .WithColumnParser("price_100", value => new(value.Value.AsT0 * 100))
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_JsonFlatObjectWithParsers.xlsx");
    }
}
