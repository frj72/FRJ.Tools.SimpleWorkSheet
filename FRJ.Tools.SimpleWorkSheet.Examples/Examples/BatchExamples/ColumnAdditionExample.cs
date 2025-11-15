using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BatchExamples;

public class ColumnAdditionExample : IExample
{
    public string Name => "Column Addition";
    public string Description => "Adding multiple cells in a column";

    private static readonly string[] SourceArrayMonths = ["Jan", "Feb", "Mar", "Apr", "May"];
    private static readonly int[] SourceArraySales= [100, 150, 200, 175, 225];

    public void Run()
    {
        var sheet = new WorkSheet("ColumnAddition");
        
        var months = SourceArrayMonths.Select(m => new CellValue(m));
        
        sheet.AddColumn(0, 0, months, cell => cell
            .WithFont(font => font.Bold()));
        
        var sales = SourceArraySales.Select(s => new CellValue(s));
        
        sheet.AddColumn(1, 0, sales);
        
        ExampleRunner.SaveWorkSheet(sheet, "11_ColumnAddition.xlsx");
    }
}