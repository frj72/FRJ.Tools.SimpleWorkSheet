using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class SheetVisibilityExample : IExample
{
    public string Name => "Sheet Visibility";
    public string Description => "Hiding sheets for data and calculations";

    public int ExampleNumber { get; }

    public SheetVisibilityExample(int exampleNumber) => ExampleNumber = exampleNumber;

    public void Run()
    {
        var reportSheet = new WorkSheet("Sales Report");
        reportSheet.AddCell(new(0, 0), "Monthly Sales Report", cell => cell
            .WithFont(font => font.Bold().WithSize(16)));
        reportSheet.AddCell(new(0, 2), "Total Sales:", null);
        reportSheet.AddCell(new(1, 2), new CellFormula("=SUM(RawData!B:B)"), cell => cell
            .WithFont(font => font.Bold()));
        reportSheet.AddCell(new(0, 3), "Average Sale:", null);
        reportSheet.AddCell(new(1, 3), new CellFormula("=AVERAGE(RawData!B:B)"), cell => cell
            .WithFont(font => font.Bold()));
        reportSheet.AddCell(new(0, 4), "Total Items:", null);
        reportSheet.AddCell(new(1, 4), new CellFormula("=COUNTA(RawData!A:A)-1"), cell => cell
            .WithFont(font => font.Bold()));

        var rawDataSheet = new WorkSheet("RawData");
        rawDataSheet.AddCell(new(0, 0), "Product", null);
        rawDataSheet.AddCell(new(1, 0), "Amount", null);
        rawDataSheet.AddCell(new(0, 1), "Widget A", null);
        rawDataSheet.AddCell(new(1, 1), 1250, null);
        rawDataSheet.AddCell(new(0, 2), "Widget B", null);
        rawDataSheet.AddCell(new(1, 2), 2340, null);
        rawDataSheet.AddCell(new(0, 3), "Widget C", null);
        rawDataSheet.AddCell(new(1, 3), 890, null);
        rawDataSheet.AddCell(new(0, 4), "Widget D", null);
        rawDataSheet.AddCell(new(1, 4), 1670, null);
        rawDataSheet.AddCell(new(0, 5), "Widget E", null);
        rawDataSheet.AddCell(new(1, 5), 3120, null);
        rawDataSheet.SetVisible(false);
        rawDataSheet.SetTabColor("808080");

        var calculationsSheet = new WorkSheet("Calculations");
        calculationsSheet.AddCell(new(0, 0), "Internal Calculations", null);
        calculationsSheet.AddCell(new(0, 1), "Tax Rate:", null);
        calculationsSheet.AddCell(new(1, 1), 0.08m, null);
        calculationsSheet.AddCell(new(0, 2), "Discount:", null);
        calculationsSheet.AddCell(new(1, 2), 0.15m, null);
        calculationsSheet.SetVisible(false);
        calculationsSheet.SetTabColor("404040");

        var workbook = new WorkBook("SalesReport", [reportSheet, rawDataSheet, calculationsSheet]);
        
        var outputPath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "Output",
            $"{ExampleNumber:000}_SheetVisibility.xlsx");

        workbook.SaveToFile(outputPath);

        Console.WriteLine($"Saved: {outputPath}");
        Console.WriteLine("✓ Completed successfully");
        Console.WriteLine();
        Console.WriteLine("Note: Open the file in Excel to see that 'RawData' and 'Calculations' sheets are hidden.");
        Console.WriteLine("      You can unhide them in Excel via: Right-click sheet tab → Unhide");
    }
}
