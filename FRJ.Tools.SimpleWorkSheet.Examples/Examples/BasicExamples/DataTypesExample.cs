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
        
        sheet.AddCell(0, 0, "String Value", null);
        sheet.AddCell(0, 1, "Integer:", null);
        sheet.AddCell(1, 1, 42, null);
        sheet.AddCell(0, 2, "Decimal:", null);
        sheet.AddCell(1, 2, 3.14159m, null);
        sheet.AddCell(0, 3, "Date:", null);
        sheet.AddCell(1, 3, DateTime.Now, null);
        sheet.AddCell(0, 4, "DateTimeOffset:", null);
        sheet.AddCell(1, 4, DateTimeOffset.Now, null);
        
        ExampleRunner.SaveWorkSheet(sheet, "002_DataTypes.xlsx");
    }
}