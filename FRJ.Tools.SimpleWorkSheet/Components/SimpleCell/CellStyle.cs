namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public record CellStyle
{
    public string? FillColor { get; init; }
    public CellFont? Font { get; init; }
    public CellBorders? Borders { get; init; }
    public string? FormatCode { get; init; }
    public HorizontalAlignment? HorizontalAlignment { get; init; }
    public VerticalAlignment? VerticalAlignment { get; init; }
    public int? TextRotation { get; init; }
    public bool? WrapText { get; init; }

    public static CellStyle Create(string? fillColor = null, CellFont? font = null, CellBorders? borders = null,
        string? formatCode = null, HorizontalAlignment? horizontalAlignment = null,
        VerticalAlignment? verticalAlignment = null, int? textRotation = null, bool? wrapText = null)
        => new()
        {
            FillColor = fillColor,
            Font = font,
            Borders = borders,
            FormatCode = formatCode,
            HorizontalAlignment = horizontalAlignment,
            VerticalAlignment = verticalAlignment,
            TextRotation = textRotation,
            WrapText = wrapText
        };
}