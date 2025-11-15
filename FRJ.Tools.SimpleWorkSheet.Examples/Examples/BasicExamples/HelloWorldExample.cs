using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

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