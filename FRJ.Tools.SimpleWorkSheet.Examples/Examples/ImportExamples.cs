using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples;

public class MetadataTrackingExample : IExample
{
    public string Name => "Metadata Tracking";
    public string Description => "Adding cells with source and import metadata";

    public void Run()
    {
        var sheet = new WorkSheet("MetadataTracking");
        
        sheet.AddCell(0, 0, "Manual Entry", cell => cell
            .WithMetadata(meta => meta
                .WithSource("manual")
                .WithImportedAt(DateTime.UtcNow)));
        
        sheet.AddCell(0, 1, "API Import", cell => cell
            .WithMetadata(meta => meta
                .WithSource("api")
                .WithImportedAt(DateTime.UtcNow)
                .AddCustomData("api_endpoint", "/users/123")
                .AddCustomData("version", "v2")));
        
        sheet.AddCell(0, 2, "Calculated", cell => cell
            .WithMetadata(meta => meta
                .WithSource("calculation")
                .AddCustomData("formula", "A1 + A2")));
        
        ExampleRunner.SaveWorkSheet(sheet, "13_MetadataTracking.xlsx");
    }
}

public class ImportOptionsExample : IExample
{
    public string Name => "Import Options";
    public string Description => "Configuring ImportOptions for different scenarios";

    public void Run()
    {
        var sheet = new WorkSheet("ImportOptions");
        
        var basicOptions = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("csv")
            .WithTrimWhitespace(true)
            .WithPreserveOriginalValue(true)
            .Build();
        
        var rawValue = "  Sample Data  ";
        var processedValue = rawValue.ProcessRawValue(basicOptions);
        
        sheet.AddCell(0, 0, processedValue, cell => cell
            .FromImportedValue(rawValue, basicOptions));
        
        var advancedOptions = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("json")
            .WithDefaultStyle(style => style
                .WithFillColor("E0E0E0")
                .WithFont(font => font.WithSize(11)))
            .WithColumnParser(0, s => new CellValue(int.Parse(s)))
            .WithCustomMetadata("import_batch", "batch_001")
            .Build();
        
        sheet.AddCell(0, 1, "Advanced Import", cell => cell
            .FromImportedValue("raw_json_value", advancedOptions));
        
        ExampleRunner.SaveWorkSheet(sheet, "14_ImportOptions.xlsx");
    }
}

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
            .WithColumnParser(1, s => new CellValue(int.Parse(s)))
            .Build();
        
        for (uint row = 0; row < csvLines.Length; row++)
        {
            var values = csvLines[row].Split(',');
            for (uint col = 0; col < values.Length; col++)
            {
                var processedValue = values[col].ProcessRawValue(options);
                
                if (row == 0)
                    sheet.AddCell(col, row, processedValue, cell => cell
                        .WithColor("4472C4")
                        .WithFont(font => font.Bold().WithColor("FFFFFF"))
                        .FromImportedValue(values[col], options));
                else
                    sheet.AddCell(col, row, processedValue, cell => cell
                        .FromImportedValue(values[col], options));
            }
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "15_CsvSimulation.xlsx");
    }
}
