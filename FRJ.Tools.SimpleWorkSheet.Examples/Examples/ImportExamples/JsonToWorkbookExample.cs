using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonToWorkbookExample : IExample
{
    public string Name => "JSON to Workbook";
    public string Description => "Creates a workbook from JSON with single data sheet";

    public void Run()
    {
        var jsonPath = Path.Combine("Resources", "Data", "Json", "Prices.json");
        
        var workbook = WorkbookBuilder.FromJsonFile(jsonPath)
            .WithWorkbookName("Price Analysis")
            .WithDataSheetName("Price Data")
            .WithPreserveOriginalValue(true)
            .WithHeaderStyle(style => style
                .WithFillColor("4472C4")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithColumnParser("Price", value => new(value.Value.AsT0 * 100))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, "61_JsonToWorkbook.xlsx");
    }
}
