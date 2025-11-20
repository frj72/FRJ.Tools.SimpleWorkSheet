using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.VisualInspection;

public class GradebookExample : IShowcase
{
    public string Name => "Student Gradebook";
    public string Description => "Gradebook with formulas and conditional colors";
    public string Category => "Visual Inspection";

    public void Run()
    {
        var sheet = new WorkSheet("Gradebook");
        
        sheet.AddCell(new(0, 0), "Student", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(1, 0), "Quiz 1", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(2, 0), "Quiz 2", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(3, 0), "Midterm", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(4, 0), "Final", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(5, 0), "Average", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        sheet.AddCell(new(6, 0), "Grade", cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        
        var students = new[]
        {
            ("Alice Johnson", 85, 92, 88, 90),
            ("Bob Smith", 78, 82, 80, 85),
            ("Carol White", 95, 98, 96, 97),
            ("David Brown", 70, 75, 72, 78),
            ("Eve Davis", 88, 85, 90, 92)
        };
        
        uint row = 1;
        foreach (var (name, quiz1, quiz2, midterm, final) in students)
        {
            sheet.AddCell(new(0, row), name, null);
            sheet.AddCell(new(1, row), quiz1, null);
            sheet.AddCell(new(2, row), quiz2, null);
            sheet.AddCell(new(3, row), midterm, null);
            sheet.AddCell(new(4, row), final, null);
            sheet.AddCell(new(5, row), new CellFormula($"=AVERAGE(B{row + 1}:E{row + 1})"), cell => cell.WithFormatCode("0.00"));
            sheet.AddCell(new(6, row), new CellFormula($"=IF(F{row + 1}>=90,\"A\",IF(F{row + 1}>=80,\"B\",IF(F{row + 1}>=70,\"C\",\"F\")))"), null);
            row++;
        }
        
        sheet.SetColumnWidth(0, 18.0);
        for (uint col = 1; col <= 6; col++)
            sheet.SetColumnWidth(col, 10.0);
        
        sheet.FreezePanes(1, 0);
        
        var workbook = new WorkBook("Gradebook", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_15_Gradebook.xlsx");
    }
}
