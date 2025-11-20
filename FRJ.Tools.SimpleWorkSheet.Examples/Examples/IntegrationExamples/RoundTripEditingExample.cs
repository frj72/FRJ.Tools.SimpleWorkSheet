using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.IntegrationExamples;

public class RoundTripEditingExample : IExample
{
    public string Name => "Round-Trip Editing";
    public string Description => "Load an Excel file, modify it, and save it back";

    public void Run()
    {
        var originalSheet = new WorkSheet("Employees");
        
        originalSheet.AddCell(new(0, 0), "ID", cell => cell.WithFont(font => font.Bold()));
        originalSheet.AddCell(new(1, 0), "Name", cell => cell.WithFont(font => font.Bold()));
        originalSheet.AddCell(new(2, 0), "Salary", cell => cell.WithFont(font => font.Bold()));
        originalSheet.AddCell(new(3, 0), "Status", cell => cell.WithFont(font => font.Bold()));
        
        originalSheet.AddCell(new(0, 1), 101, null);
        originalSheet.AddCell(new(1, 1), "Alice", null);
        originalSheet.AddCell(new(2, 1), 75000, null);
        originalSheet.AddCell(new(3, 1), "Active", null);
        
        originalSheet.AddCell(new(0, 2), 102, null);
        originalSheet.AddCell(new(1, 2), "Bob", null);
        originalSheet.AddCell(new(2, 2), 68000, null);
        originalSheet.AddCell(new(3, 2), "Active", null);
        
        var tempPath = Path.Combine(Path.GetTempPath(), "roundtrip_temp.xlsx");
        ExampleRunner.SaveWorkSheet(originalSheet, tempPath);
        
        Console.WriteLine("Step 1: Created original file with 2 employees");
        Console.WriteLine();
        
        var loadedWorkbook = WorkBookReader.LoadFromFile(tempPath);
        var loadedSheet = loadedWorkbook.Sheets.First();
        
        Console.WriteLine($"Step 2: Loaded file - found {loadedSheet.Cells.Cells.Count} cells");
        Console.WriteLine();
        
        loadedSheet.AddCell(new(0, 3), 103, null);
        loadedSheet.AddCell(new(1, 3), "Charlie", null);
        loadedSheet.AddCell(new(2, 3), 82000, null);
        loadedSheet.AddCell(new(3, 3), "Active", cell => cell
            .WithColor("90EE90"));
        
        var existingCell = loadedSheet.Cells.Cells[new(3, 1)];
        loadedSheet.Cells.Cells[new(3, 1)] = existingCell with 
        { 
            Style = existingCell.Style with { FillColor = "FFD700" }
        };
        
        Console.WriteLine("Step 3: Added new employee (Charlie) and highlighted Alice's status");
        Console.WriteLine();
        
        ExampleRunner.SaveWorkSheet(loadedSheet, "36_RoundTripEditing.xlsx");
        
        Console.WriteLine("Step 4: Saved modified file");
        Console.WriteLine($"  Total cells now: {loadedSheet.Cells.Cells.Count}");
        Console.WriteLine("  New employee added at row 3");
        Console.WriteLine("  Status cell highlighted for employee 101");
        
        File.Delete(tempPath);
    }
}
