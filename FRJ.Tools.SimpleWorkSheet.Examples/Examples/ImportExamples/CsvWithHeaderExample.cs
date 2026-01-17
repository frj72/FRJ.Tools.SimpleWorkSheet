using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class CsvWithHeaderExample : IExample
{
    public string Name => "CSV Import with Header";
    public string Description => "Imports CSV file with header row using fluent API";


    public int ExampleNumber { get; }

    public CsvWithHeaderExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var csvPath = Path.Combine("Resources", "Data", "Csv", "MixedTypes.csv");
        
        var sheet = WorksheetBuilder.FromCsvFile(csvPath, hasHeader: true)
            .WithSheetName("CSV Data")
            .WithHeaderStyle(style => style
                .WithFillColor("70AD47")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_CsvWithHeader.xlsx");
    }
}
