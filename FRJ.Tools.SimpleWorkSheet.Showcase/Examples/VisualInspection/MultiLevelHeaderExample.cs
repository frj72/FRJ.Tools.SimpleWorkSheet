using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.VisualInspection;

public class MultiLevelHeaderExample : IShowcase
{
    public string Name => "Complex Table with Nested Headers";
    public string Description => "Multi-level header structure with merged cells";
    public string Category => "Visual Inspection";

    public void Run()
    {
        var sheet = new WorkSheet("ComplexTable");
        
        sheet.AddCell(new(0, 0), "Product Sales Analysis - 2025", cell => cell
            .WithFont(f => f.Bold().WithSize(14))
            .WithStyle(s => s.WithHorizontalAlignment(HorizontalAlignment.Center)));
        sheet.MergeCells(0, 0, 9, 0);
        
        sheet.AddCell(new(0, 1), "", cell => cell.WithColor("4472C4"));
        sheet.AddCell(new(1, 1), "Q1", cell => cell
            .WithFont(f => f.Bold())
            .WithColor("4472C4")
            .WithStyle(s => s.WithHorizontalAlignment(HorizontalAlignment.Center)));
        sheet.MergeCells(1, 1, 3, 1);
        sheet.AddCell(new(4, 1), "Q2", cell => cell
            .WithFont(f => f.Bold())
            .WithColor("4472C4")
            .WithStyle(s => s.WithHorizontalAlignment(HorizontalAlignment.Center)));
        sheet.MergeCells(4, 1, 6, 1);
        sheet.AddCell(new(7, 1), "Q3", cell => cell
            .WithFont(f => f.Bold())
            .WithColor("4472C4")
            .WithStyle(s => s.WithHorizontalAlignment(HorizontalAlignment.Center)));
        sheet.MergeCells(7, 1, 9, 1);
        
        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep" };
        sheet.AddCell(new(0, 2), "Product", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        for (uint i = 0; i < 9; i++)
            sheet.AddCell(new(i + 1, 2), months[i], cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        
        var products = new[] { "Widget A", "Widget B", "Widget C" };
        for (uint row = 0; row < 3; row++)
        {
            sheet.AddCell(new(0, row + 3), products[row], null);
            for (uint col = 1; col <= 9; col++)
                sheet.AddCell(new(col, row + 3), (row + 1) * col * 100, null);
        }
        
        for (uint col = 0; col <= 9; col++)
            sheet.SetColumnWidth(col, 10.0);
        
        sheet.FreezePanes(1, 3);
        
        var workbook = new WorkBook("MultiLevel", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_18_MultiLevelHeader.xlsx");
    }
}
