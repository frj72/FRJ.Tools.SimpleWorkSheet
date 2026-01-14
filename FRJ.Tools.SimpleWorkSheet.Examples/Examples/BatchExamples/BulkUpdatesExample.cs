using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BatchExamples;

public class BulkUpdatesExample : IExample
{
    public string Name => "Bulk Updates";
    public string Description => "Using UpdateCell for modifying existing data";

    public void Run()
    {
        var sheet = new WorkSheet("BulkUpdates");
        
        for (uint i = 0; i < 5; i++) 
            sheet.AddCell(0, i, $"Original {i}", null);

        for (uint i = 0; i < 5; i++)
        {
            var i1 = i;
            sheet.UpdateCell(0, i, cell => cell
                .WithValue($"Updated {i1}")
                .WithColor("FFFF00")
                .WithFont(font => font.Italic()));
        }

        ExampleRunner.SaveWorkSheet(sheet, "012_BulkUpdates.xlsx");
    }
}