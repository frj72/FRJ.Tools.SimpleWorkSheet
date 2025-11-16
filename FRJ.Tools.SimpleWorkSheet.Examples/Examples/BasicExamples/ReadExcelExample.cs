using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class ReadExcelExample : IExample
{
    public string Name => "Read Excel File";
    public string Description => "Demonstrates loading and reading existing Excel files";

    public void Run()
    {
        var originalSheet = new WorkSheet("SampleData");
        
        originalSheet.AddCell(new(0, 0), "Name", cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        originalSheet.AddCell(new(1, 0), "Age", cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        originalSheet.AddCell(new(2, 0), "Email", cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        
        originalSheet.AddCell(new(0, 1), "Alice");
        originalSheet.AddCell(new(1, 1), 30);
        originalSheet.AddCell(new(2, 1), "alice@example.com", cell => cell
            .WithHyperlink("mailto:alice@example.com"));
        
        originalSheet.AddCell(new(0, 2), "Bob");
        originalSheet.AddCell(new(1, 2), 25);
        originalSheet.AddCell(new(2, 2), "bob@example.com", cell => cell
            .WithHyperlink("mailto:bob@example.com"));
        
        originalSheet.SetColumnWith(0, 15.0);
        originalSheet.SetColumnWith(2, 25.0);
        originalSheet.FreezePanes(1, 0);
        
        var tempPath = Path.Combine(Path.GetTempPath(), "read_example_temp.xlsx");
        ExampleRunner.SaveWorkSheet(originalSheet, tempPath);
        
        Console.WriteLine("Created temporary Excel file for demonstration");
        Console.WriteLine($"File location: {tempPath}");
        Console.WriteLine();
        
        var loadedWorkbook = WorkBookReader.LoadFromFile(tempPath);
        
        Console.WriteLine($"Loaded workbook with {loadedWorkbook.Sheets.Count()} sheet(s)");
        Console.WriteLine();
        
        foreach (var sheet in loadedWorkbook.Sheets)
        {
            Console.WriteLine($"Sheet: {sheet.Name}");
            Console.WriteLine($"  Cells: {sheet.Cells.Cells.Count}");
            Console.WriteLine($"  Frozen panes: {(sheet.FrozenPane != null ? $"Row={sheet.FrozenPane.Row}, Col={sheet.FrozenPane.Column}" : "None")}");
            Console.WriteLine($"  Column widths: {sheet.ExplicitColumnWidths.Count}");
            Console.WriteLine();
            
            Console.WriteLine("Cell Values:");
            foreach (var cellEntry in sheet.Cells.Cells.Take(6))
            {
                var pos = cellEntry.Key;
                var cell = cellEntry.Value;
                var valueStr = cell.Value.IsString() ? cell.Value.ToString() : 
                              cell.Value.IsLong() ? cell.Value.ToString() :
                              cell.Value.IsDecimal() ? cell.Value.ToString() : "?";
                
                Console.WriteLine($"  [{pos.X},{pos.Y}]: {valueStr}");
            }
        }
        
        var loadedSheet = loadedWorkbook.Sheets.First();
        ExampleRunner.SaveWorkSheet(loadedSheet, "35_ReadExcel.xlsx");
        
        File.Delete(tempPath);
    }
}
