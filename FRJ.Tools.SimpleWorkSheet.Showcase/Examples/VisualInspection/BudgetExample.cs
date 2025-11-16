using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.VisualInspection;

public class BudgetExample : IShowcase
{
    public string Name => "Budget Template";
    public string Description => "Monthly budget with named ranges";
    public string Category => "Visual Inspection";

    public void Run()
    {
        var sheet = new WorkSheet("Budget");
        
        sheet.AddCell(new(0, 0), "Monthly Budget", cell => cell.WithFont(f => f.Bold().WithSize(14)));
        
        sheet.AddCell(new(0, 2), "Income", cell => cell.WithFont(f => f.Bold()).WithColor("A9D08E"));
        sheet.AddCell(new(0, 3), "Salary");
        sheet.AddCell(new(1, 3), 5000m, cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(0, 4), "Freelance");
        sheet.AddCell(new(1, 4), 1000m, cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(0, 5), "Total Income", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 5), new CellFormula("=SUM(B4:B5)"), cell => cell
            .WithFormatCode("$#,##0.00")
            .WithFont(f => f.Bold())
            .WithColor("D4F4DD"));
        
        sheet.AddCell(new(0, 7), "Expenses", cell => cell.WithFont(f => f.Bold()).WithColor("F4B084"));
        sheet.AddCell(new(0, 8), "Rent");
        sheet.AddCell(new(1, 8), 1500m, cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(0, 9), "Food");
        sheet.AddCell(new(1, 9), 600m, cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(0, 10), "Transportation");
        sheet.AddCell(new(1, 10), 300m, cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(0, 11), "Utilities");
        sheet.AddCell(new(1, 11), 200m, cell => cell.WithFormatCode("$#,##0.00"));
        sheet.AddCell(new(0, 12), "Total Expenses", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 12), new CellFormula("=SUM(B9:B12)"), cell => cell
            .WithFormatCode("$#,##0.00")
            .WithFont(f => f.Bold())
            .WithColor("F8CBAD"));
        
        sheet.AddCell(new(0, 14), "Net Savings", cell => cell.WithFont(f => f.Bold().WithSize(12)));
        sheet.AddCell(new(1, 14), new CellFormula("=B6-B13"), cell => cell
            .WithFormatCode("$#,##0.00")
            .WithFont(f => f.Bold().WithSize(12))
            .WithColor("FFD966"));
        
        sheet.SetColumnWith(0, 20.0);
        sheet.SetColumnWith(1, 15.0);
        
        var workbook = new WorkBook("Budget", [sheet]);
        workbook.AddNamedRange("TotalIncome", "Budget", 1, 5, 1, 5);
        workbook.AddNamedRange("TotalExpenses", "Budget", 1, 12, 1, 12);
        workbook.AddNamedRange("NetSavings", "Budget", 1, 14, 1, 14);
        
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_17_Budget.xlsx");
    }
}
