using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonArrayWithStylingExample : IExample
{
    public string Name => "JSON Array Import with Styling";
    public string Description => "Imports JSON array with styled headers (color, bold font) and auto-fitted columns";


    public int ExampleNumber { get; }

    public JsonArrayWithStylingExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices.json");
        
        var sheet = WorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Prices with Styling")
            .WithHeaderStyle(style => style
                .WithFillColor("4472C4")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_JsonArrayWithStyling.xlsx");
    }
}
