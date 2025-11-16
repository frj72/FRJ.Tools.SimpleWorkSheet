using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellValidationTests
{
    [Fact]
    public void ListValidation_CreatesCorrectFormula()
    {
        var validation = CellValidation.List(["Option A", "Option B", "Option C"]);

        Assert.Equal(ValidationType.List, validation.Type);
        Assert.Equal("\"Option A,Option B,Option C\"", validation.Formula1);
        Assert.True(validation.AllowBlank);
    }

    [Fact]
    public void WholeNumberValidation_Between_SetsCorrectValues()
    {
        var validation = CellValidation.WholeNumber(ValidationOperator.Between, 1, 100);

        Assert.Equal(ValidationType.WholeNumber, validation.Type);
        Assert.Equal(ValidationOperator.Between, validation.Operator);
        Assert.Equal("1", validation.Formula1);
        Assert.Equal("100", validation.Formula2);
    }

    [Fact]
    public void DecimalValidation_GreaterThan_SetsCorrectValues()
    {
        var validation = CellValidation.DecimalNumber(ValidationOperator.GreaterThan, 10.5);

        Assert.Equal(ValidationType.Decimal, validation.Type);
        Assert.Equal(ValidationOperator.GreaterThan, validation.Operator);
        Assert.Equal("10.5", validation.Formula1);
        Assert.Null(validation.Formula2);
    }

    [Fact]
    public void DateValidation_Between_UsesOADate()
    {
        var date1 = new DateTime(2025, 1, 1);
        var date2 = new DateTime(2025, 12, 31);
        var validation = CellValidation.Date(ValidationOperator.Between, date1, date2);

        Assert.Equal(ValidationType.Date, validation.Type);
        Assert.Equal(date1.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), validation.Formula1);
        Assert.Equal(date2.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), validation.Formula2);
    }

    [Fact]
    public void TextLengthValidation_LessThan_SetsCorrectValues()
    {
        var validation = CellValidation.TextLength(ValidationOperator.LessThan, 50);

        Assert.Equal(ValidationType.TextLength, validation.Type);
        Assert.Equal(ValidationOperator.LessThan, validation.Operator);
        Assert.Equal("50", validation.Formula1);
    }

    [Fact]
    public void TimeValidation_Record_CanBeConstructed()
    {
        var validation = new CellValidation
        {
            Type = ValidationType.Time,
            Operator = ValidationOperator.Between,
            Formula1 = "0.25",
            Formula2 = "0.75"
        };

        Assert.Equal(ValidationType.Time, validation.Type);
        Assert.Equal("0.25", validation.Formula1);
        Assert.Equal("0.75", validation.Formula2);
    }

    [Fact]
    public void CustomValidation_SetsFormula()
    {
        var validation = CellValidation.Custom("=A1>B1");

        Assert.Equal(ValidationType.Custom, validation.Type);
        Assert.Equal("=A1>B1", validation.Formula1);
    }

    [Fact]
    public void WholeNumberValidation_NotBetween_SetsOperator()
    {
        var validation = CellValidation.WholeNumber(ValidationOperator.NotBetween, 5, 10);

        Assert.Equal(ValidationOperator.NotBetween, validation.Operator);
        Assert.Equal("5", validation.Formula1);
        Assert.Equal("10", validation.Formula2);
    }

    [Fact]
    public void DecimalValidation_Equal_SetsOperator()
    {
        var validation = CellValidation.DecimalNumber(ValidationOperator.Equal, 12.5);

        Assert.Equal(ValidationOperator.Equal, validation.Operator);
        Assert.Equal("12.5", validation.Formula1);
        Assert.Null(validation.Formula2);
    }

    [Fact]
    public void DecimalValidation_NotEqual_SetsOperator()
    {
        var validation = CellValidation.DecimalNumber(ValidationOperator.NotEqual, 3.14);

        Assert.Equal(ValidationOperator.NotEqual, validation.Operator);
        Assert.Equal("3.14", validation.Formula1);
    }

    [Fact]
    public void WholeNumberValidation_GreaterThanOrEqual_SetsOperator()
    {
        var validation = CellValidation.WholeNumber(ValidationOperator.GreaterThanOrEqual, 100);

        Assert.Equal(ValidationOperator.GreaterThanOrEqual, validation.Operator);
        Assert.Equal("100", validation.Formula1);
    }

    [Fact]
    public void WholeNumberValidation_LessThanOrEqual_SetsOperator()
    {
        var validation = CellValidation.WholeNumber(ValidationOperator.LessThanOrEqual, 50);

        Assert.Equal(ValidationOperator.LessThanOrEqual, validation.Operator);
        Assert.Equal("50", validation.Formula1);
    }

    [Fact]
    public void WithInputMessage_SetsProperties()
    {
        var validation = CellValidation.List(["A", "B"])
            .WithInputMessage("Title", "Message");

        Assert.True(validation.ShowInputMessage);
        Assert.Equal("Title", validation.InputTitle);
        Assert.Equal("Message", validation.InputMessage);
    }

    [Fact]
    public void WithErrorAlert_SetsProperties()
    {
        var validation = CellValidation.List(["A", "B"])
            .WithErrorAlert("Error Title", "Error Message", "warning");

        Assert.True(validation.ShowErrorAlert);
        Assert.Equal("Error Title", validation.ErrorTitle);
        Assert.Equal("Error Message", validation.ErrorMessage);
        Assert.Equal("warning", validation.ErrorStyle);
    }

    [Fact]
    public void AddValidation_ToWorkSheet_AddsToValidationsDictionary()
    {
        var sheet = new WorkSheet("Test");
        var validation = CellValidation.List(["A", "B", "C"]);

        sheet.AddValidation(0, 0, validation);

        Assert.Single(sheet.Validations);
        Assert.True(sheet.Validations.ContainsKey(CellRange.FromPositions(new(0, 0), new(0, 0))));
    }

    [Fact]
    public void AddValidation_WithRange_AddsToValidationsDictionary()
    {
        var sheet = new WorkSheet("Test");
        var validation = CellValidation.WholeNumber(ValidationOperator.Between, 1, 100);
        var range = CellRange.FromBounds(0, 0, 5, 10);

        sheet.AddValidation(range, validation);

        Assert.Single(sheet.Validations);
        Assert.Contains(range, sheet.Validations.Keys);
    }

    [Fact]
    public void AddValidation_WithCoordinates_AddsToValidationsDictionary()
    {
        var sheet = new WorkSheet("Test");
        var validation = CellValidation.DecimalNumber(ValidationOperator.GreaterThan, 0.0);

        sheet.AddValidation(1, 1, 5, 5, validation);

        Assert.Single(sheet.Validations);
        var expectedRange = CellRange.FromBounds(1, 1, 5, 5);
        Assert.Contains(expectedRange, sheet.Validations.Keys);
    }

    [Fact]
    public void ListValidation_WithAllowBlankFalse_SetsCorrectly()
    {
        var validation = CellValidation.List(["A", "B"], allowBlank: false);

        Assert.False(validation.AllowBlank);
    }
}
