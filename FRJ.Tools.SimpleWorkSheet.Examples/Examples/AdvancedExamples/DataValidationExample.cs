using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class DataValidationExample : IExample
{
    public string Name => "Data Validation";
    public string Description => "Demonstrates various data validation rules";

    public void Run()
    {
        var sheet = new WorkSheet("Validation");

        sheet.AddCell(new(0, 0), "Validation Type", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("DDDDDD")));
        sheet.AddCell(new(1, 0), "Input Cell", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("DDDDDD")));
        sheet.AddCell(new(2, 0), "Description", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("DDDDDD")));

        sheet.AddCell(new(0, 2), "Dropdown List");
        sheet.AddCell(new(1, 2), "Select one:");
        sheet.AddCell(new(2, 2), "Options: Red, Green, Blue");
        var listValidation = CellValidation.List(["Red", "Green", "Blue"])
            .WithInputMessage("Color Selection", "Please select a color from the list");
        sheet.AddValidation(1, 2, listValidation);

        sheet.AddCell(new(0, 4), "Whole Number");
        sheet.AddCell(new(1, 4), "Enter number:");
        sheet.AddCell(new(2, 4), "Must be between 1 and 100");
        var numberValidation = CellValidation.WholeNumber(ValidationOperator.Between, 1, 100)
            .WithInputMessage("Number Input", "Enter a whole number between 1 and 100")
            .WithErrorAlert("Invalid Number", "Number must be between 1 and 100", "stop");
        sheet.AddValidation(1, 4, numberValidation);

        sheet.AddCell(new(0, 6), "Decimal Number");
        sheet.AddCell(new(1, 6), "Enter decimal:");
        sheet.AddCell(new(2, 6), "Must be greater than 0");
        var decimalValidation = CellValidation.DecimalNumber(ValidationOperator.GreaterThan, 0.0)
            .WithInputMessage("Decimal Input", "Enter a positive decimal number");
        sheet.AddValidation(1, 6, decimalValidation);

        sheet.AddCell(new(0, 8), "Date Range");
        sheet.AddCell(new(1, 8), "Enter date:");
        sheet.AddCell(new(2, 8), "Must be in 2025");
        var dateValidation = CellValidation.Date(
            ValidationOperator.Between,
            new DateTime(2025, 1, 1),
            new DateTime(2025, 12, 31))
            .WithInputMessage("Date Input", "Enter a date in 2025")
            .WithErrorAlert("Invalid Date", "Date must be between 01/01/2025 and 12/31/2025", "warning");
        sheet.AddValidation(1, 8, dateValidation);

        sheet.AddCell(new(0, 10), "Text Length");
        sheet.AddCell(new(1, 10), "Enter text:");
        sheet.AddCell(new(2, 10), "Maximum 50 characters");
        var textLengthValidation = CellValidation.TextLength(ValidationOperator.LessThanOrEqual, 50)
            .WithInputMessage("Text Input", "Enter up to 50 characters");
        sheet.AddValidation(1, 10, textLengthValidation);

        sheet.AddCell(new(0, 12), "Custom Formula");
        sheet.AddCell(new(1, 12), "Enter value:");
        sheet.AddCell(new(2, 12), "Must be even number");
        var customValidation = CellValidation.Custom("=MOD(B13,2)=0")
            .WithInputMessage("Even Number", "Enter an even number")
            .WithErrorAlert("Invalid Input", "Value must be an even number", "information");
        sheet.AddValidation(1, 12, customValidation);

        sheet.AddCell(new(0, 14), "Range Validation");
        sheet.AddCell(new(1, 14), "A");
        sheet.AddCell(new(2, 14), "B");
        sheet.AddCell(new(3, 14), "C");
        sheet.AddCell(new(4, 14), "D");
        sheet.AddCell(new(2, 15), "Multiple cells with same validation");
        var rangeValidation = CellValidation.WholeNumber(ValidationOperator.Between, 0, 10);
        sheet.AddValidation(1, 14, 4, 14, rangeValidation);

        sheet.SetColumnWith(0, 20.0);
        sheet.SetColumnWith(1, 15.0);
        sheet.SetColumnWith(2, 40.0);

        ExampleRunner.SaveWorkSheet(sheet, "38_DataValidation.xlsx");
    }
}
