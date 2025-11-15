using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class CsvSimulationExample : IExample
{
    public string Name => "CSV Simulation";
    public string Description => "Simulating CSV import with metadata preservation";

    public void Run()
    {
        var sheet = new WorkSheet("CsvSimulation");
        
        var csvLines = new[]
        {
            "Name,Age,City",
            "John Doe,30,NYC",
            "Jane Smith,25,LA",
            "Bob Johnson,35,Chicago"
        };
        
        var options = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("csv")
            .WithTrimWhitespace(true)
            .WithPreserveOriginalValue(true)
            .WithColumnParser(1, s => new(int.Parse(s)))
            .Build();
        
        for (uint row = 0; row < csvLines.Length; row++)
        {
            var values = csvLines[row].Split(',');
            for (uint col = 0; col < values.Length; col++)
            {
                var processedValue = values[col].ProcessRawValue(options);
                
                if (row == 0)
                {
                    var col1 = col;
                    sheet.AddCell(col, row, processedValue, cell => cell
                        .WithColor("4472C4")
                        .WithFont(font => font.Bold().WithColor("FFFFFF"))
                        .FromImportedValue(values[col1], options));
                }
                else
                {
                    var col1 = col;
                    sheet.AddCell(col, row, processedValue, cell => cell
                        .FromImportedValue(values[col1], options));
                }
            }
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "15_CsvSimulation.xlsx");
    }
}