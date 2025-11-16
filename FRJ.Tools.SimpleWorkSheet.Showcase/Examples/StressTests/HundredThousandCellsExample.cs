using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.StressTests;

public class HundredThousandCellsExample : IShowcase
{
    public string Name => "100,000 Cells with Styling";
    public string Description => "Stress test with 100,000 styled cells";
    public string Category => "Stress Tests";

    public void Run()
    {
        var sheet = new WorkSheet("StressTest");
        
        Console.WriteLine("  Creating 100,000 cells...");
        var cellCount = 0;
        
        for (uint row = 0; row < 500; row++)
        for (uint col = 0; col < 200; col++)
        {
            var isEven = (row + col) % 2 == 0;
            sheet.AddCell(new(col, row), $"R{row}C{col}", cell =>
            {
                if (isEven)
                    cell.WithColor("E8F4F8");
                cell.WithFont(f => f.WithSize(9));
            });
            cellCount++;
                
            if (cellCount % 10000 == 0)
                Console.WriteLine($"    Created {cellCount} cells...");
        }

        var workbook = new WorkBook("100KCells", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_09_HundredThousandCells.xlsx");
    }
}
