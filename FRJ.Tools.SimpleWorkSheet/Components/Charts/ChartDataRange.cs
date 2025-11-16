using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public static class ChartDataRange
{
    public static void ValidateDataRange(CellRange range)
    {
        if (range.From.X == range.To.X && range.From.Y == range.To.Y)
            throw new ArgumentException("Data range must contain more than one cell", nameof(range));
    }

    public static string ToRangeReference(CellRange range, string sheetName)
    {
        var fromCol = GetColumnName(range.From.X + 1);
        var fromRow = range.From.Y + 1;
        var toCol = GetColumnName(range.To.X + 1);
        var toRow = range.To.Y + 1;
        return $"'{sheetName}'!${fromCol}${fromRow}:${toCol}${toRow}";
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
