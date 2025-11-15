using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

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