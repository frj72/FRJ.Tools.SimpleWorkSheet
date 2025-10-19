using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples;

public class HelloWorldExample : IExample
{
    public string Name => "Hello World";
    public string Description => "Simple cell creation and basic worksheet setup";

    public void Run()
    {
        var sheet = new WorkSheet("HelloWorld");
        sheet.AddCell(0, 0, "Hello World");
        ExampleRunner.SaveWorkSheet(sheet, "01_HelloWorld.xlsx");
    }
}

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

public class CellPositioningExample : IExample
{
    public string Name => "Cell Positioning";
    public string Description => "Using coordinates vs CellPosition objects";

    public void Run()
    {
        var sheet = new WorkSheet("Positioning");
        
        sheet.AddCell(0, 0, "Using x, y coordinates");
        
        var position = new CellPosition(0, 1);
        sheet.AddCell(position, "Using CellPosition object");
        
        sheet.AddCell(new CellPosition(0, 2), "Another CellPosition");
        
        ExampleRunner.SaveWorkSheet(sheet, "04_CellPositioning.xlsx");
    }
}

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
