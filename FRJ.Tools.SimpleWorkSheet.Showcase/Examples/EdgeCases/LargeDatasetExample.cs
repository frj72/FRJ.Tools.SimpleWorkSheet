using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.EdgeCases;

public class LargeDatasetExample : IShowcase
{
    public string Name => "Large Dataset (10,000 rows × 20 columns)";
    public string Description => "Tests performance and stability with large data volumes";
    public string Category => "Edge Cases";

    public void Run()
    {
        var sheet = new WorkSheet("LargeData");
        
        for (uint col = 0; col < 20; col++)
            sheet.AddCell(new(col, 0), $"Column{col + 1}", cell => cell
                .WithFont(f => f.Bold().WithSize(12))
                .WithColor("4472C4"));
        
        for (uint row = 1; row <= 10000; row++)
        {
            for (uint col = 0; col < 20; col++)
            {
                var value = col switch
                {
                    0 => new(row),
                    1 => new($"Row {row}"),
                    2 => new(row * 1.5m),
                    3 => new(DateTime.Now.AddDays(row)),
                    _ => new CellValue($"Data{row}-{col}")
                };
                sheet.AddCell(new(col, row), value, null);
            }
            
            if (row % 1000 == 0)
                Console.WriteLine($"  Generated {row} rows...");
        }
        
        sheet.FreezePanes(1, 0);
        
        var workbook = new WorkBook("LargeDataset", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_01_LargeDataset.xlsx");
    }
}
