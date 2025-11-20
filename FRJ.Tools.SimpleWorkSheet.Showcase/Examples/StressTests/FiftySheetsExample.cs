using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.StressTests;

public class FiftySheetsExample : IShowcase
{
    public string Name => "50 Sheets in One Workbook";
    public string Description => "Tests workbook with many sheets";
    public string Category => "Stress Tests";

    public void Run()
    {
        var sheets = new List<WorkSheet>();
        
        Console.WriteLine("  Creating 50 sheets...");
        for (var i = 1; i <= 50; i++)
        {
            var sheet = new WorkSheet($"Sheet{i}");
            
            for (uint row = 0; row < 100; row++)
                for (uint col = 0; col < 10; col++)
                    sheet.AddCell(new(col, row), $"S{i}-R{row}-C{col}", null);
            
            sheets.Add(sheet);
            
            if (i % 10 == 0)
                Console.WriteLine($"    Created {i} sheets...");
        }
        
        var workbook = new WorkBook("FiftySheets", [.. sheets]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_10_FiftySheets.xlsx");
    }
}
