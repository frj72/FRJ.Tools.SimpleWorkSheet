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
    public CellHyperlink? Hyperlink { get; init; }

    public string? Color => Style.FillColor;
    public CellFont? Font => Style.Font;
    public CellBorders? Borders => Style.Borders;
    public string? FormatCode => Style.FormatCode;
}