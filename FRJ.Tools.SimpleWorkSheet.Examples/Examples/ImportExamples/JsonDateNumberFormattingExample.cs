using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonDateNumberFormattingExample : IExample
{
    public string Name => "JSON with Date and Number Formatting";
    public string Description => "Applies consistent date formatting and number formatting to imported JSON data";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices.json");
        
        var sheet = JsonWorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Formatted Prices")
            .WithDateFormat(DateFormat.DateOnly)
            .WithNumberFormat("Price", NumberFormat.Float2)
            .WithHeaderStyle(style => style
                .WithFillColor("5B9BD5")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "67_JsonDateNumberFormatting.xlsx");
    }
}
