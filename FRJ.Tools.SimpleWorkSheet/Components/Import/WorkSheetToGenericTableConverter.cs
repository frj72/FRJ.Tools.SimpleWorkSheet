using System.Globalization;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public static class WorkSheetToGenericTableConverter
{
    public static GenericTable Convert(WorkSheet sheet, HeaderMode headerMode = HeaderMode.None)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        var bounds = GetDataBounds(sheet);
        if (bounds == null)
            return new GenericTable();

        var (minX, minY, maxX, maxY) = bounds.Value;

        var (headers, dataStartRow) = ExtractHeaders(sheet, headerMode, minX, minY, maxX, maxY);

        var table = new GenericTable(headers);

        for (var row = dataStartRow; row <= maxY; row++)
        {
            var rowValues = new List<CellValue?>();

            for (var col = minX; col <= maxX; col++)
            {
                var position = new CellPosition(col, row);
                var value = sheet.Cells.Cells.TryGetValue(position, out var cell)
                    ? ExtractCellValue(cell.Value)
                    : null;

                rowValues.Add(value);
            }

            table.AddRow(rowValues);
        }

        return table;
    }

    public static Dictionary<string, GenericTable> ConvertWorkBook(WorkBook workbook, HeaderMode headerMode = HeaderMode.None)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        var result = new Dictionary<string, GenericTable>();

        foreach (var sheet in workbook.Sheets)
        {
            var table = Convert(sheet, headerMode);

            if (table is { RowCount: 0, ColumnCount: 0 })
                continue;

            result[sheet.Name] = table;
        }

        return result;
    }

    private static (uint minX, uint minY, uint maxX, uint maxY)? GetDataBounds(WorkSheet sheet)
    {
        if (sheet.Cells.Cells.Count == 0)
            return null;

        var positions = sheet.Cells.Cells.Keys;

        var minX = positions.Min(p => p.X);
        var minY = positions.Min(p => p.Y);
        var maxX = positions.Max(p => p.X);
        var maxY = positions.Max(p => p.Y);

        return (minX, minY, maxX, maxY);
    }

    private static (List<string> headers, uint dataStartRow) ExtractHeaders(
        WorkSheet sheet,
        HeaderMode headerMode,
        uint minX,
        uint minY,
        uint maxX,
        uint maxY)
    {
        var columnCount = (int)(maxX - minX + 1);

        return headerMode switch
        {
            HeaderMode.FirstRow => ExtractFirstRowHeaders(sheet, minX, minY, maxX),
            HeaderMode.AutoDetect => DetectAndExtractHeaders(sheet, minX, minY, maxX, maxY),
            _ => (GenerateColumnNames(columnCount), minY)
        };
    }

    private static (List<string> headers, uint dataStartRow) ExtractFirstRowHeaders(
        WorkSheet sheet,
        uint minX,
        uint minY,
        uint maxX)
    {
        var headers = new List<string>();

        for (var col = minX; col <= maxX; col++)
        {
            var position = new CellPosition(col, minY);
            var headerValue = sheet.Cells.Cells.TryGetValue(position, out var cell)
                ? ExtractStringFromCellValue(cell.Value)
                : $"Column{col - minX + 1}";

            headers.Add(headerValue);
        }

        return (headers, minY + 1);
    }

    private static string ExtractStringFromCellValue(CellValue cellValue)
    {
        var union = cellValue.Value;

        if (union.IsT0)
            return union.AsT0.ToString(CultureInfo.InvariantCulture);
        if (union.IsT1)
            return union.AsT1.ToString(CultureInfo.InvariantCulture);
        if (union.IsT2)
            return union.AsT2 ?? "Column1";
        if (union.IsT3)
            return union.AsT3.ToString(CultureInfo.InvariantCulture);

        return union.IsT4
            ? union.AsT4.ToString(CultureInfo.InvariantCulture)
            : union.AsT5.Value;
    }

    private static (List<string> headers, uint dataStartRow) DetectAndExtractHeaders(
        WorkSheet sheet,
        uint minX,
        uint minY,
        uint maxX,
        uint maxY)
    {
        var isLikelyHeader = true;

        for (var col = minX; col <= maxX; col++)
        {
            var position = new CellPosition(col, minY);
            if (sheet.Cells.Cells.TryGetValue(position, out var cell) && cell.Value.IsString())
                continue;

            isLikelyHeader = false;
            break;
        }

        if (!isLikelyHeader || minY >= maxY)
        {
            var columnCount = (int)(maxX - minX + 1);
            return (GenerateColumnNames(columnCount), minY);
        }

        var hasDifferentTypesInSecondRow = false;
        for (var col = minX; col <= maxX; col++)
        {
            var position = new CellPosition(col, minY + 1);
            if (!sheet.Cells.Cells.TryGetValue(position, out var cell))
                continue;

            if (!cell.Value.IsDecimal() && !cell.Value.IsLong())
                continue;

            hasDifferentTypesInSecondRow = true;
            break;
        }

        if (hasDifferentTypesInSecondRow)
            return ExtractFirstRowHeaders(sheet, minX, minY, maxX);

        var colCount = (int)(maxX - minX + 1);
        return (GenerateColumnNames(colCount), minY);
    }

    private static List<string> GenerateColumnNames(int count)
    {
        var names = new List<string>();
        for (var i = 1; i <= count; i++)
            names.Add($"Column{i}");

        return names;
    }

    private static CellValue ExtractCellValue(CellValue cellValue)
    {
        var union = cellValue.Value;

        if (union.IsT5)
            return new CellValue(union.AsT5.Value);
        if (union.IsT0)
            return new CellValue(union.AsT0);
        if (union.IsT1)
            return new CellValue(union.AsT1);
        if (union.IsT2)
            return new CellValue(union.AsT2 ?? string.Empty);

        return union.IsT3
            ? new CellValue(union.AsT3)
            : new CellValue(union.AsT4);
    }
}
