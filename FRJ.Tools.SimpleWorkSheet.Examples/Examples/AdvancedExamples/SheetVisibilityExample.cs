using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class SheetVisibilityExample : IExample
{
    public string Name => "Sheet Visibility";
    public string Description => "Hiding sheets for data and calculations";

    public void Run()
    {
        var reportSheet = new WorkSheet("Sales Report");
        reportSheet.AddCell(new(0, 0), "Monthly Sales Report", cell => cell
            .WithFont(font => font.Bold().WithSize(16)));
        reportSheet.AddCell(new(0, 2), "Total Sales:");
        reportSheet.AddCell(new(1, 2), new CellFormula("=SUM(RawData!B:B)"), cell => cell
            .WithFont(font => font.Bold()));
        reportSheet.AddCell(new(0, 3), "Average Sale:");
        reportSheet.AddCell(new(1, 3), new CellFormula("=AVERAGE(RawData!B:B)"), cell => cell
            .WithFont(font => font.Bold()));
        reportSheet.AddCell(new(0, 4), "Total Items:");
        reportSheet.AddCell(new(1, 4), new CellFormula("=COUNTA(RawData!A:A)-1"), cell => cell
            .WithFont(font => font.Bold()));

        var rawDataSheet = new WorkSheet("RawData");
        rawDataSheet.AddCell(new(0, 0), "Product");
        rawDataSheet.AddCell(new(1, 0), "Amount");
        rawDataSheet.AddCell(new(0, 1), "Widget A");
        rawDataSheet.AddCell(new(1, 1), 1250);
        rawDataSheet.AddCell(new(0, 2), "Widget B");
        rawDataSheet.AddCell(new(1, 2), 2340);
        rawDataSheet.AddCell(new(0, 3), "Widget C");
        rawDataSheet.AddCell(new(1, 3), 890);
        rawDataSheet.AddCell(new(0, 4), "Widget D");
        rawDataSheet.AddCell(new(1, 4), 1670);
        rawDataSheet.AddCell(new(0, 5), "Widget E");
        rawDataSheet.AddCell(new(1, 5), 3120);
        rawDataSheet.SetVisible(false);
        rawDataSheet.SetTabColor("808080");

        var calculationsSheet = new WorkSheet("Calculations");
        calculationsSheet.AddCell(new(0, 0), "Internal Calculations");
        calculationsSheet.AddCell(new(0, 1), "Tax Rate:");
        calculationsSheet.AddCell(new(1, 1), 0.08m);
        calculationsSheet.AddCell(new(0, 2), "Discount:");
        calculationsSheet.AddCell(new(1, 2), 0.15m);
        calculationsSheet.SetVisible(false);
        calculationsSheet.SetTabColor("404040");

        var workbook = new WorkBook("SalesReport", [reportSheet, rawDataSheet, calculationsSheet]);
        
        var outputPath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "Output",
            "48_SheetVisibility.xlsx");

        workbook.SaveToFile(outputPath);

        Console.WriteLine($"Saved: {outputPath}");
        Console.WriteLine("✓ Completed successfully");
        Console.WriteLine();
        Console.WriteLine("Note: Open the file in Excel to see that 'RawData' and 'Calculations' sheets are hidden.");
        Console.WriteLine("      You can unhide them in Excel via: Right-click sheet tab → Unhide");
    }
}
