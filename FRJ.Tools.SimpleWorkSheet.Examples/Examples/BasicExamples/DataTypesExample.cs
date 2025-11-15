using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class DataTypesExample : IExample
{
    public string Name => "Data Types";
    public string Description => "Demonstrating different value types";

    public void Run()
    {
        var sheet = new WorkSheet("DataTypes");
        
        sheet.AddCell(0, 0, "String Value");
        sheet.AddCell(0, 1, "Integer:");
        sheet.AddCell(1, 1, 42);
        sheet.AddCell(0, 2, "Decimal:");
        sheet.AddCell(1, 2, 3.14159m);
        sheet.AddCell(0, 3, "Date:");
        sheet.AddCell(1, 3, DateTime.Now);
        sheet.AddCell(0, 4, "DateTimeOffset:");
        sheet.AddCell(1, 4, DateTimeOffset.Now);
        
        ExampleRunner.SaveWorkSheet(sheet, "02_DataTypes.xlsx");
    }
}