namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public record CellValidation
{
    public ValidationType Type { get; init; }
    public ValidationOperator Operator { get; init; }
    public string? Formula1 { get; init; }
    public string? Formula2 { get; init; }
    public bool AllowBlank { get; init; } = true;
    public bool ShowInputMessage { get; init; }
    public string? InputTitle { get; init; }
    public string? InputMessage { get; init; }
    public bool ShowErrorAlert { get; init; } = true;
    public string? ErrorTitle { get; init; }
    public string? ErrorMessage { get; init; }
    public string? ErrorStyle { get; init; } = "stop";

    public static CellValidation List(IEnumerable<string> items, bool allowBlank = true)
    {
        var formula = $"\"{string.Join(",", items)}\"";
        return new()
        {
            Type = ValidationType.List,
            Operator = ValidationOperator.Between,
            Formula1 = formula,
            AllowBlank = allowBlank
        };
    }

    public static CellValidation WholeNumber(ValidationOperator op, int value1, int? value2 = null, bool allowBlank = true)
    {
        return new()
        {
            Type = ValidationType.WholeNumber,
            Operator = op,
            Formula1 = value1.ToString(),
            Formula2 = value2?.ToString(),
            AllowBlank = allowBlank
        };
    }

    public static CellValidation DecimalNumber(ValidationOperator op, double value1, double? value2 = null, bool allowBlank = true)
    {
        return new()
        {
            Type = ValidationType.Decimal,
            Operator = op,
            Formula1 = value1.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Formula2 = value2?.ToString(System.Globalization.CultureInfo.InvariantCulture),
            AllowBlank = allowBlank
        };
    }

    public static CellValidation Date(ValidationOperator op, DateTime value1, DateTime? value2 = null, bool allowBlank = true)
    {
        return new()
        {
            Type = ValidationType.Date,
            Operator = op,
            Formula1 = value1.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture),
            Formula2 = value2?.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture),
            AllowBlank = allowBlank
        };
    }

    public static CellValidation TextLength(ValidationOperator op, int length1, int? length2 = null, bool allowBlank = true)
    {
        return new()
        {
            Type = ValidationType.TextLength,
            Operator = op,
            Formula1 = length1.ToString(),
            Formula2 = length2?.ToString(),
            AllowBlank = allowBlank
        };
    }

    public static CellValidation Custom(string formula, bool allowBlank = true)
    {
        return new()
        {
            Type = ValidationType.Custom,
            Operator = ValidationOperator.Between,
            Formula1 = formula,
            AllowBlank = allowBlank
        };
    }

    public CellValidation WithInputMessage(string title, string message)
    {
        return this with
        {
            ShowInputMessage = true,
            InputTitle = title,
            InputMessage = message
        };
    }

    public CellValidation WithErrorAlert(string title, string message, string style = "stop")
    {
        return this with
        {
            ShowErrorAlert = true,
            ErrorTitle = title,
            ErrorMessage = message,
            ErrorStyle = style
        };
    }
}
