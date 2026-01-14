using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class SheetTabColorsExample : IExample
{
    public string Name => "Sheet Tab Colors";
    public string Description => "Setting custom colors for sheet tabs";

    public void Run()
    {
        var summarySheet = new WorkSheet("Summary");
        summarySheet.AddCell(new(0, 0), "Quarterly Report Summary", cell => cell
            .WithFont(font => font.Bold().WithSize(16))
            .WithStyle(style => style.WithFillColor("4472C4")));
        summarySheet.AddCell(new(0, 2), "Q1: $125,000", null);
        summarySheet.AddCell(new(0, 3), "Q2: $142,000", null);
        summarySheet.AddCell(new(0, 4), "Q3: $138,500", null);
        summarySheet.AddCell(new(0, 5), "Q4: $159,200", null);
        summarySheet.SetTabColor("4472C4");

        var q1Sheet = new WorkSheet("Q1 Sales");
        q1Sheet.AddCell(new(0, 0), "Q1 Sales Data", cell => cell
            .WithFont(font => font.Bold()));
        q1Sheet.AddCell(new(0, 1), "Product A", null);
        q1Sheet.AddCell(new(1, 1), 45000, null);
        q1Sheet.AddCell(new(0, 2), "Product B", null);
        q1Sheet.AddCell(new(1, 2), 80000, null);
        q1Sheet.SetTabColor("92D050");

        var q2Sheet = new WorkSheet("Q2 Sales");
        q2Sheet.AddCell(new(0, 0), "Q2 Sales Data", cell => cell
            .WithFont(font => font.Bold()));
        q2Sheet.AddCell(new(0, 1), "Product A", null);
        q2Sheet.AddCell(new(1, 1), 52000, null);
        q2Sheet.AddCell(new(0, 2), "Product B", null);
        q2Sheet.AddCell(new(1, 2), 90000, null);
        q2Sheet.SetTabColor("00B050");

        var q3Sheet = new WorkSheet("Q3 Sales");
        q3Sheet.AddCell(new(0, 0), "Q3 Sales Data", cell => cell
            .WithFont(font => font.Bold()));
        q3Sheet.AddCell(new(0, 1), "Product A", null);
        q3Sheet.AddCell(new(1, 1), 48500, null);
        q3Sheet.AddCell(new(0, 2), "Product B", null);
        q3Sheet.AddCell(new(1, 2), 90000, null);
        q3Sheet.SetTabColor("FFC000");

        var q4Sheet = new WorkSheet("Q4 Sales");
        q4Sheet.AddCell(new(0, 0), "Q4 Sales Data", cell => cell
            .WithFont(font => font.Bold()));
        q4Sheet.AddCell(new(0, 1), "Product A", null);
        q4Sheet.AddCell(new(1, 1), 59200, null);
        q4Sheet.AddCell(new(0, 2), "Product B", null);
        q4Sheet.AddCell(new(1, 2), 100000, null);
        q4Sheet.SetTabColor("FF0000");

        var workbook = new WorkBook("QuarterlyReport", [summarySheet, q1Sheet, q2Sheet, q3Sheet, q4Sheet]);
        
        var outputPath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "Output",
            "047_SheetTabColors.xlsx");

        workbook.SaveToFile(outputPath);

        Console.WriteLine($"Saved: {outputPath}");
        Console.WriteLine("✓ Completed successfully");
    }
}
