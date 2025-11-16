using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class NamedRangesExample : IExample
{
    public string Name => "Named Ranges";
    public string Description => "Demonstrates named ranges and formulas using them";

    public void Run()
    {
        var sheet = new WorkSheet("Sales");

        sheet.AddCell(new(0, 0), "Product", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Q1", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(2, 0), "Q2", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(3, 0), "Q3", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(4, 0), "Q4", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(5, 0), "Total", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.AddCell(new(0, 1), "Product A");
        sheet.AddCell(new(1, 1), 1250.0);
        sheet.AddCell(new(2, 1), 1420.0);
        sheet.AddCell(new(3, 1), 1380.0);
        sheet.AddCell(new(4, 1), 1550.0);

        sheet.AddCell(new(0, 2), "Product B");
        sheet.AddCell(new(1, 2), 980.0);
        sheet.AddCell(new(2, 2), 1150.0);
        sheet.AddCell(new(3, 2), 1200.0);
        sheet.AddCell(new(4, 2), 1320.0);

        sheet.AddCell(new(0, 3), "Product C");
        sheet.AddCell(new(1, 3), 1820.0);
        sheet.AddCell(new(2, 3), 1950.0);
        sheet.AddCell(new(3, 3), 2100.0);
        sheet.AddCell(new(4, 3), 2250.0);

        sheet.AddCell(new(5, 1), new CellFormula("=SUM(B2:E2)"));
        sheet.AddCell(new(5, 2), new CellFormula("=SUM(B3:E3)"));
        sheet.AddCell(new(5, 3), new CellFormula("=SUM(B4:E4)"));

        sheet.AddCell(new(0, 5), "Summary Statistics", cell => cell
            .WithFont(font => font.Bold().WithSize(12))
            .WithStyle(style => style.WithFillColor("DDDDDD")));

        sheet.AddCell(new(0, 7), "Total All Products:");
        sheet.AddCell(new(1, 7), new CellFormula("=SUM(TotalSales)"));

        sheet.AddCell(new(0, 8), "Average Per Product:");
        sheet.AddCell(new(1, 8), new CellFormula("=AVERAGE(TotalSales)"));

        sheet.AddCell(new(0, 9), "Max Quarter (All):");
        sheet.AddCell(new(1, 9), new CellFormula("=MAX(QuarterlySales)"));

        sheet.AddCell(new(0, 10), "Min Quarter (All):");
        sheet.AddCell(new(1, 10), new CellFormula("=MIN(QuarterlySales)"));

        sheet.SetColumnWith(0, 20.0);
        sheet.SetColumnWith(5, 15.0);

        var workbook = new WorkBook("Named Ranges Demo", [sheet]);

        workbook.AddNamedRange("QuarterlySales", "Sales", 1, 1, 4, 3);
        workbook.AddNamedRange("TotalSales", "Sales", 5, 1, 5, 3);

        var bytes = FRJ.Tools.SimpleWorkSheet.LowLevel.SheetConverter.ToBinaryExcelFile(workbook);
        var outputPath = Path.Combine("Output", "39_NamedRanges.xlsx");
        File.WriteAllBytes(outputPath, bytes);
    }
}
