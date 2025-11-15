using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class FreezePanesExample : IExample
{
    public string Name => "Freeze Panes";
    public string Description => "Demonstrates freezing rows and columns";

    public void Run()
    {
        var sheet = new WorkSheet("FreezePanes");
        
        var headers = new[] { "Column A", "Column B", "Column C", "Column D", "Column E" }
            .Select(h => new CellValue(h));
        sheet.AddRow(0, 0, headers, cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        
        for (uint row = 1; row < 20; row++)
        {
            for (uint col = 0; col < 5; col++)
            {
                sheet.AddCell(new CellPosition(col, row), $"R{row}C{col}");
            }
        }
        
        sheet.FreezePanes(1, 0);
        
        ExampleRunner.SaveWorkSheet(sheet, "32_FreezePanes.xlsx");
    }
}
