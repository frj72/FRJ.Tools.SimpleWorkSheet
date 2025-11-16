using FRJ.Tools.SimpleWorkSheet.Components.Book;

namespace FRJ.Tools.SimpleWorkSheet.Showcase;

public static class ShowcaseRunner
{
    public static void SaveWorkBook(WorkBook workBook, string filename)
    {
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        
        var filePath = Path.Combine(outputDir, filename);
        workBook.SaveToFile(filePath);
        
        Console.WriteLine($"Saved: {filePath}");
    }

    public static void RunShowcase(IShowcase test)
    {
        Console.WriteLine($"\n=== {test.Category}: {test.Name} ===");
        Console.WriteLine(test.Description);
        Console.WriteLine();
        
        try
        {
            test.Run();
            Console.WriteLine("✓ Completed successfully");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
