using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.EdgeCases;

public class ExtremeSizesExample : IShowcase
{
    public string Name => "Extreme Sizes (Column widths, Row heights)";
    public string Description => "Tests maximum column widths and row heights";
    public string Category => "Edge Cases";

    public void Run()
    {
        var sheet = new WorkSheet("ExtremeSizes");
        
        sheet.AddCell(new(0, 0), "Description", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 0), "Value", cell => cell.WithFont(f => f.Bold()));
        
        sheet.AddCell(new(0, 1), "Minimum column width (1.0)");
        sheet.AddCell(new(1, 1), "Tiny");
        sheet.SetColumnWith(1, 1.0);
        
        sheet.AddCell(new(0, 3), "Maximum column width (255.0)");
        sheet.AddCell(new(2, 3), "This column is very wide - Excel maximum");
        sheet.SetColumnWith(2, 255.0);
        
        sheet.AddCell(new(0, 5), "Minimum row height (1.0)");
        sheet.AddCell(new(1, 5), "Short");
        sheet.SetRowHeight(5, 1.0);
        
        sheet.AddCell(new(0, 7), "Maximum row height (409.0)");
        sheet.AddCell(new(1, 7), "Tall row");
        sheet.SetRowHeight(7, 409.0);
        
        sheet.AddCell(new(0, 9), "Normal size for comparison");
        sheet.AddCell(new(1, 9), "Default");
        
        sheet.SetColumnWith(0, 30.0);
        
        var workbook = new WorkBook("ExtremeSizes", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_08_ExtremeSizes.xlsx");
    }
}
