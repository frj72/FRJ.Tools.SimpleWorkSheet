using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BatchExamples;

public class RowAdditionExample : IExample
{
    public string Name => "Row Addition";
    public string Description => "Adding multiple cells in a row with consistent styling";

    private static readonly string[] SourceArrayHeaders = ["Product", "Price", "Quantity", "Total"];
    private static readonly string[] SourceArrayRow1 = ["Widget", "9.99", "10", "99.90"];
    private static readonly string[] SourceArrayRow2 = ["Gadget", "19.99", "5", "99.95"];

    public void Run()
    {
        var sheet = new WorkSheet("RowAddition");
        
        var headers = SourceArrayHeaders.Select(h => new CellValue(h));
        
        sheet.AddRow(0, 0, headers, configure: cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithColor("FFFFFF")));
        
        var row1 = SourceArrayRow1.Select(v => new CellValue(v));
        
        sheet.AddRow(1, 0, row1, configure: cell => cell.WithColor("F0F0F0"));
        
        var row2 = SourceArrayRow2.Select(v => new CellValue(v));
        
        sheet.AddRow(2, 0, row2, null);
        
        ExampleRunner.SaveWorkSheet(sheet, "10_RowAddition.xlsx");
    }
}