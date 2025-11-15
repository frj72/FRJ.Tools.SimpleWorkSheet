using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class ValueOnlyExample : IExample
{
    public string Name => "Value Only";
    public string Description => "Adding cells with just values, no styling";

    public void Run()
    {
        var sheet = new WorkSheet("ValueOnly");
        
        for (uint i = 0; i < 5; i++)
        {
            sheet.AddCell(0, i, $"Row {i + 1}");
            sheet.AddCell(1, i, (i + 1) * 10);
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "05_ValueOnly.xlsx");
    }
}