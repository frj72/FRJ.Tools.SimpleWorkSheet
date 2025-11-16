using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.EdgeCases;

public class MaximumColumnsExample : IShowcase
{
    public string Name => "Maximum Columns (Excel's 16,384 limit)";
    public string Description => "Tests Excel's maximum column limit (XFD column)";
    public string Category => "Edge Cases";

    public void Run()
    {
        var sheet = new WorkSheet("MaxColumns");
        
        Console.WriteLine("  Creating 16,384 columns...");
        for (uint col = 0; col < 16384; col++)
        {
            sheet.AddCell(new(col, 0), $"Col{col + 1}");
            sheet.AddCell(new(col, 1), col);
            
            if ((col + 1) % 1000 == 0)
                Console.WriteLine($"    Created {col + 1} columns...");
        }
        
        var workbook = new WorkBook("MaxColumns", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_02_MaximumColumns.xlsx");
    }
}
