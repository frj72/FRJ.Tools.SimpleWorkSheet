using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class ValueOnlyExample : IExample
{
    public string Name => "Value Only";
    public string Description => "Adding cells with just values, no styling";


    public int ExampleNumber { get; }

    public ValueOnlyExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("ValueOnly");
        
        for (uint i = 0; i < 5; i++)
        {
            sheet.AddCell(0, i, $"Row {i + 1}", null);
            sheet.AddCell(1, i, (i + 1) * 10, null);
        }
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_ValueOnly.xlsx");
    }
}