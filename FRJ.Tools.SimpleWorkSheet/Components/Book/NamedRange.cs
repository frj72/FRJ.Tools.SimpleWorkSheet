using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Book;

public record NamedRange
{
    public string Name { get; init; }
    public string SheetName { get; init; }
    public CellRange Range { get; init; }

    public NamedRange(string name, string sheetName, CellRange range)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Name cannot be null or whitespace", nameof(name));
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet name cannot be null or whitespace", nameof(sheetName));
        
        Name = name;
        SheetName = sheetName;
        Range = range;
    }

    public string ToFormulaReference()
    {
        var fromCol = GetColumnName(Range.From.X + 1);
        var fromRow = Range.From.Y + 1;
        var toCol = GetColumnName(Range.To.X + 1);
        var toRow = Range.To.Y + 1;
        return $"'{SheetName}'!${fromCol}${fromRow}:${toCol}${toRow}";
    }

    private static string GetColumnName(uint columnNumber)
    {
        var dividend = columnNumber;
        var columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }
}
