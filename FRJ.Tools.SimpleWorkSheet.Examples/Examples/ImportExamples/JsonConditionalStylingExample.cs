using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonConditionalStylingExample : IExample
{
    public string Name => "JSON with Conditional Styling";
    public string Description => "Applies conditional styling based on cell values in imported JSON data";


    public int ExampleNumber { get; }

    public JsonConditionalStylingExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Orders.json");
        
        var sheet = WorksheetBuilder.FromJsonFile(jsonPath)
            .WithSheetName("Orders with Styling")
            .WithHeaderStyle(style => style
                .WithFillColor("ED7D31")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithConditionalStyle("status",
                val => val.IsString() && val.Value.AsT2 == "pending",
                style => style.WithFillColor("FFF2CC"))
            .WithConditionalStyle("totalAmount",
                val => val.IsDecimal() && val.Value.AsT0 > 500,
                style => style.WithFillColor("C6E0B4"))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_JsonConditionalStyling.xlsx");
    }
}
