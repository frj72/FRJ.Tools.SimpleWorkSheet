using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ValidationShowcase;

public class FormInputExample : IShowcase
{
    public string Name => "Form-like Interface with Dropdowns";
    public string Description => "User input form with validation";
    public string Category => "Validation Showcase";

    public void Run()
    {
        var sheet = new WorkSheet("Form");
        
        sheet.AddCell(new(0, 0), "Customer Information Form", cell => cell
            .WithFont(f => f.Bold().WithSize(14))
            .WithColor("4472C4"));
        sheet.MergeCells(0, 0, 1, 0);
        
        sheet.AddCell(new(0, 2), "Name:", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 2), "[Enter name]");
        
        sheet.AddCell(new(0, 3), "Email:", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 3), "[Enter email]");
        
        sheet.AddCell(new(0, 4), "Country:", cell => cell.WithFont(f => f.Bold()));
        var countryValidation = CellValidation.List(["USA", "Canada", "UK", "Australia", "Other"]);
        sheet.AddValidation(1, 4, countryValidation);
        
        sheet.AddCell(new(0, 5), "Age:", cell => cell.WithFont(f => f.Bold()));
        var ageValidation = CellValidation.WholeNumber(ValidationOperator.Between, 18, 100)
            .WithInputMessage("Age", "Please enter age between 18-100")
            .WithErrorAlert("Invalid Age", "Age must be between 18 and 100");
        sheet.AddValidation(1, 5, ageValidation);
        
        sheet.AddCell(new(0, 6), "Subscription:", cell => cell.WithFont(f => f.Bold()));
        var subValidation = CellValidation.List(["Basic", "Premium", "Enterprise"]);
        sheet.AddValidation(1, 6, subValidation);
        
        sheet.AddCell(new(0, 7), "Start Date:", cell => cell.WithFont(f => f.Bold()));
        var dateValidation = CellValidation.Date(ValidationOperator.GreaterThanOrEqual, DateTime.Today);
        sheet.AddValidation(1, 7, dateValidation);
        
        sheet.SetColumnWith(0, 15.0);
        sheet.SetColumnWith(1, 25.0);
        
        var workbook = new WorkBook("FormInput", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_23_FormInput.xlsx");
    }
}
