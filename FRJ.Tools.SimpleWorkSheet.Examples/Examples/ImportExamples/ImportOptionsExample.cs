using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

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
        
        const string rawValue = "  Sample Data  ";
        var processedValue = rawValue.ProcessRawValue(basicOptions);
        
        sheet.AddCell(0, 0, processedValue, cell => cell
            .FromImportedValue(rawValue, basicOptions));
        
        var advancedOptions = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("json")
            .WithDefaultStyle(style => style
                .WithFillColor("E0E0E0")
                .WithFont(font => font.WithSize(11)))
            .WithColumnParser(0, s => new(int.Parse(s)))
            .WithCustomMetadata("import_batch", "batch_001")
            .Build();
        
        sheet.AddCell(0, 1, "Advanced Import", cell => cell
            .FromImportedValue("raw_json_value", advancedOptions));
        
        ExampleRunner.SaveWorkSheet(sheet, "14_ImportOptions.xlsx");
    }
}