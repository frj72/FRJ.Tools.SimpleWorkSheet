using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class FreezePanesExample : IExample
{
    public string Name => "Freeze Panes";
    public string Description => "Demonstrates freezing rows and columns";


    public int ExampleNumber { get; }

    public FreezePanesExample(int exampleNumber) => ExampleNumber = exampleNumber;
    private static readonly string[] SourceArrayHeaders = ["Column A", "Column B", "Column C", "Column D", "Column E"];

    public void Run()
    {
        var sheet = new WorkSheet("FreezePanes");
        
        var headers = SourceArrayHeaders.Select(h => new CellValue(h));
        sheet.AddRow(0, 0, headers, configure: cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        
        for (uint row = 1; row < 20; row++)
        for (uint col = 0; col < 5; col++) sheet.AddCell(new(col, row), $"R{row}C{col}", null);

        sheet.FreezePanes(1, 0);
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_FreezePanes.xlsx");
    }
}
