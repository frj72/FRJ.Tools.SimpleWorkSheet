using System.Globalization;
using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public record Cell
{
    public Cell(CellValue Value, CellStyle Style, CellMetadata? Metadata = null)
    {
        this.Value = Value;
        this.Style = Style;
        this.Metadata = Metadata;
    }

    public Cell(CellValue Value, string? Color, CellFont? Font, CellBorders? Borders, string? FormatCode)
    {
        this.Value = Value;
        var formatCode = FormatCode ?? Value.CellValueType() switch
        {
            CellValueBasicType.DateType => DateFormat.IsoDateTime.ToExcelFormatString(),
            CellValueBasicType.FloatingPointNumber => NumberFormat.Float3.ToExcelFormatString(),
            CellValueBasicType.IntegerNumber => NumberFormat.Integer.ToExcelFormatString(),
            _ => null
        };
        Style = CellStyle.Create(Color, Font, Borders, formatCode);
        Metadata = null;
    }

    public static Cell Create(CellValue value, string? color = null, CellFont? font = null, CellBorders? borders = null, string? formatCode = null) =>
        new(value, color ?? WorkSheetDefaults.FillColor, font ?? WorkSheetDefaults.Font,
            borders ?? WorkSheetDefaults.CellBorders, formatCode);

    public static Cell CreateEmpty() => Create(string.Empty);

    public CellValue Value { get; init; }
    public CellStyle Style { get; init; }
    public CellMetadata? Metadata { get; init; }

    public string? Color => Style.FillColor;
    public CellFont? Font => Style.Font;
    public CellBorders? Borders => Style.Borders;
    public string? FormatCode => Style.FormatCode;
};

public static class CellExtensions
{
    public static double CalculateApproximateWith(this Cell cell)
    {
        var cellFont = cell.Font ?? WorkSheetDefaults.Font;
        var fontSize = cellFont.Size ?? WorkSheetDefaults.FontSize;
        if (cell.Value.IsCellFormula())
            return -1;
        if (cell.Value.IsDateTime() || cell.Value.IsDateTimeOffset())
        {
            return GetTextWidth(fontSize, 19);
        }
        if (cell.Value.IsDecimal())
        {
            var text = cell.Value.AsDecimal().ToString(CultureInfo.InvariantCulture);
            var textb = cell.Value.AsDecimal().ToString("F3");
            return GetTextWidth(fontSize, Math.Min(text.Length, textb.Length));
        }
        if (cell.Value.IsLong())
        {
            var text = cell.Value.AsLong().ToString(CultureInfo.InvariantCulture);
            return GetTextWidth(fontSize, text.Length);
        }
        if (cell.Value.IsString())
        {
            return GetTextWidth(fontSize, cell.Value.AsString().Length);
        }

        return -1;
    }
    
    public static double GetTextWidth(int fontSize, int textLength) => fontSize / 11.0 * 7.0 * textLength / 7.5;
}
