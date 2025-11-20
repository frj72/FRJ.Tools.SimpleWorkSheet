using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class CsvWithoutHeaderExample : IExample
{
    public string Name => "CSV Import without Header";
    public string Description => "Imports CSV file without header row - auto-generates column names";

    public void Run()
    {
        var csvPath = Path.Combine("Resources", "Data", "Csv", "MixedTypesWithoutHeader.csv");
        
        var sheet = WorksheetBuilder.FromCsvFile(csvPath, hasHeader: false)
            .WithSheetName("CSV No Header")
            .WithHeaderStyle(style => style
                .WithFillColor("5B9BD5")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "71_CsvWithoutHeader.xlsx");
    }
}
