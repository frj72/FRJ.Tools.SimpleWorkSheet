using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class SimpleTableExample : IExample
{
    public string Name => "Simple Table";
    public string Description => "Creating a basic 3x3 table with headers";


    public int ExampleNumber { get; }

    public SimpleTableExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("SimpleTable");
        
        sheet.AddCell(0, 0, "Name", null);
        sheet.AddCell(1, 0, "Age", null);
        sheet.AddCell(2, 0, "City", null);
        
        sheet.AddCell(0, 1, "John", null);
        sheet.AddCell(1, 1, 30, null);
        sheet.AddCell(2, 1, "NYC", null);
        
        sheet.AddCell(0, 2, "Jane", null);
        sheet.AddCell(1, 2, 25, null);
        sheet.AddCell(2, 2, "LA", null);
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_SimpleTable.xlsx");
    }
}