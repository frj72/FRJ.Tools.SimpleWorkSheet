using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class SimpleTableExample : IExample
{
    public string Name => "Simple Table";
    public string Description => "Creating a basic 3x3 table with headers";

    public void Run()
    {
        var sheet = new WorkSheet("SimpleTable");
        
        sheet.AddCell(0, 0, "Name");
        sheet.AddCell(1, 0, "Age");
        sheet.AddCell(2, 0, "City");
        
        sheet.AddCell(0, 1, "John");
        sheet.AddCell(1, 1, 30);
        sheet.AddCell(2, 1, "NYC");
        
        sheet.AddCell(0, 2, "Jane");
        sheet.AddCell(1, 2, 25);
        sheet.AddCell(2, 2, "LA");
        
        ExampleRunner.SaveWorkSheet(sheet, "03_SimpleTable.xlsx");
    }
}