using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.VisualInspection;

public class FinancialReportExample : IShowcase
{
    public string Name => "Financial Statement";
    public string Description => "Complex financial report with formatting";
    public string Category => "Visual Inspection";

    public void Run()
    {
        var sheet = new WorkSheet("Financial");
        
        sheet.AddCell(new(0, 0), "Company Financial Statement - Q4 2025", cell => cell
            .WithFont(f => f.Bold().WithSize(14))
            .WithStyle(s => s.WithHorizontalAlignment(HorizontalAlignment.Center)));
        sheet.MergeCells(0, 0, 5, 0);
        
        sheet.AddCell(new(0, 2), "Item", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(1, 2), "Q1", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(2, 2), "Q2", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(3, 2), "Q3", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(4, 2), "Q4", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(5, 2), "Total", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        
        var revenues = new[] { 100000m, 120000m, 115000m, 130000m };
        sheet.AddCell(new(0, 3), "Revenue", null);
        for (var i = 0; i < 4; i++)
            sheet.AddCell(new((uint)(i + 1), 3), revenues[i], cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(5, 3), new CellFormula("=SUM(B4:E4)"), cell => cell.WithFormatCode("$#,##0.00").WithFont(f => f.Bold()));
        
        var expenses = new[] { 60000m, 70000m, 68000m, 75000m };
        sheet.AddCell(new(0, 4), "Expenses", null);
        for (var i = 0; i < 4; i++)
            sheet.AddCell(new((uint)(i + 1), 4), expenses[i], cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(5, 4), new CellFormula("=SUM(B5:E5)"), cell => cell.WithFormatCode("$#,##0.00").WithFont(f => f.Bold()));
        
        sheet.AddCell(new(0, 5), "Net Profit", cell => cell.WithFont(f => f.Bold()));
        for (uint i = 1; i <= 4; i++)
            sheet.AddCell(new(i, 5), new CellFormula($"={(char)('A' + i)}4-{(char)('A' + i)}5"), cell => cell
                .WithFormatCode("$#,##0.00")
                .WithFont(f => f.Bold())
                .WithColor("D4F4DD"));
        sheet.AddCell(new(5, 5), new CellFormula("=F4-F5"), cell => cell
            .WithFormatCode("$#,##0.00")
            .WithFont(f => f.Bold().WithSize(12))
            .WithColor("A9D08E"));
        
        for (uint col = 0; col <= 5; col++)
            sheet.SetColumnWidth(col, 12.0);
        
        sheet.FreezePanes(1, 3);
        
        var workbook = new WorkBook("FinancialReport", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_13_FinancialReport.xlsx");
    }
}
