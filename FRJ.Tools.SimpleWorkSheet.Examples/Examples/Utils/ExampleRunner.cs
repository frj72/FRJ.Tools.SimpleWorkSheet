using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

public static class ExampleRunner
{
    public static void SaveWorkSheet(WorkSheet sheet, string filename)
    {
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        
        var filePath = Path.Combine(outputDir, filename);
        
        var workBook = new WorkBook("Examples", [sheet]);
        workBook.SaveToFile(filePath);
        
        Console.WriteLine($"Saved: {filePath}");
    }

    public static void SaveWorkBook(WorkBook workBook, string filename)
    {
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        
        var filePath = Path.Combine(outputDir, filename);
        
        workBook.SaveToFile(filePath);
        
        Console.WriteLine($"Saved: {filePath}");
    }

    public static void RunExample(IExample example)
    {
        Console.WriteLine($"\n=== {example.Name} ===");
        Console.WriteLine(example.Description);
        Console.WriteLine();
        
        try
        {
            example.Run();
            Console.WriteLine("✓ Completed successfully");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error: {ex.Message}");
        }
    }
}
