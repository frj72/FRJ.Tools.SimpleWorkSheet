using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ValidationShowcase;

public class DateRangeValidationExample : IShowcase
{
    public string Name => "Date Range Validation";
    public string Description => "Complex date constraints and validations";
    public string Category => "Validation Showcase";

    public void Run()
    {
        var sheet = new WorkSheet("DateValidation");
        
        sheet.AddCell(new(0, 0), "Date Constraint", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 0), "Description", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(2, 0), "Enter Date", cell => cell.WithFont(f => f.Bold()));
        
        uint row = 1;
        
        sheet.AddCell(new(0, row), "Future Dates Only");
        sheet.AddCell(new(1, row), "Date must be after today");
        var futureValidation = CellValidation.Date(ValidationOperator.GreaterThan, DateTime.Today);
        sheet.AddValidation(2, row, futureValidation);
        row++;
        
        sheet.AddCell(new(0, row), "2025 Only");
        sheet.AddCell(new(1, row), "Date must be in year 2025");
        var year2025Validation = CellValidation.Date(ValidationOperator.Between,
            new(2025, 1, 1), new DateTime(2025, 12, 31));
        sheet.AddValidation(2, row, year2025Validation);
        row++;
        
        sheet.AddCell(new(0, row), "Past Dates");
        sheet.AddCell(new(1, row), "Date must be before today");
        var pastValidation = CellValidation.Date(ValidationOperator.LessThan, DateTime.Today);
        sheet.AddValidation(2, row, pastValidation);
        row++;
        
        sheet.AddCell(new(0, row), "Next 30 Days");
        sheet.AddCell(new(1, row), "Date within next 30 days");
        var next30Validation = CellValidation.Date(ValidationOperator.Between,
            DateTime.Today, DateTime.Today.AddDays(30));
        sheet.AddValidation(2, row, next30Validation);
        
        for (uint col = 0; col < 3; col++)
            sheet.SetColumnWith(col, 25.0);
        
        var workbook = new WorkBook("DateValidation", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_24_DateRangeValidation.xlsx");
    }
}
