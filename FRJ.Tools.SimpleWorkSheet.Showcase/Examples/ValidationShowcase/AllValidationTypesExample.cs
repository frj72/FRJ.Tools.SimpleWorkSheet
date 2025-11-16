using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ValidationShowcase;

public class AllValidationTypesExample : IShowcase
{
    public string Name => "All Validation Types";
    public string Description => "Every validation type in one sheet";
    public string Category => "Validation Showcase";

    public void Run()
    {
        var sheet = new WorkSheet("AllValidations");
        
        sheet.AddCell(new(0, 0), "Validation Type", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 0), "Instructions", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(2, 0), "Try It", cell => cell.WithFont(f => f.Bold()));
        
        uint row = 1;
        
        sheet.AddCell(new(0, row), "List");
        sheet.AddCell(new(1, row), "Choose: Red, Green, or Blue");
        var listValidation = CellValidation.List(["Red", "Green", "Blue"]);
        sheet.AddValidation(2, row, listValidation);
        row++;
        
        sheet.AddCell(new(0, row), "Whole Number");
        sheet.AddCell(new(1, row), "Enter number between 1-100");
        var wholeValidation = CellValidation.WholeNumber(ValidationOperator.Between, 1, 100);
        sheet.AddValidation(2, row, wholeValidation);
        row++;
        
        sheet.AddCell(new(0, row), "Decimal");
        sheet.AddCell(new(1, row), "Enter decimal > 0");
        var decimalValidation = CellValidation.DecimalNumber(ValidationOperator.GreaterThan, 0.0);
        sheet.AddValidation(2, row, decimalValidation);
        row++;
        
        sheet.AddCell(new(0, row), "Date");
        sheet.AddCell(new(1, row), "Enter date in 2025");
        var dateValidation = CellValidation.Date(ValidationOperator.Between, 
            new(2025, 1, 1), new DateTime(2025, 12, 31));
        sheet.AddValidation(2, row, dateValidation);
        row++;
        
        sheet.AddCell(new(0, row), "Text Length");
        sheet.AddCell(new(1, row), "Enter text (max 10 chars)");
        var textValidation = CellValidation.TextLength(ValidationOperator.LessThanOrEqual, 10);
        sheet.AddValidation(2, row, textValidation);
        row++;
        
        sheet.AddCell(new(0, row), "Custom Formula");
        sheet.AddCell(new(1, row), "Enter even number");
        var customValidation = CellValidation.Custom("=MOD(C7,2)=0");
        sheet.AddValidation(2, row, customValidation);
        
        for (uint col = 0; col < 3; col++)
            sheet.SetColumnWith(col, 25.0);
        
        var workbook = new WorkBook("AllValidations", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_22_AllValidationTypes.xlsx");
    }
}
