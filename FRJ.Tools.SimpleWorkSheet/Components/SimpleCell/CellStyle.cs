namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public record CellStyle
{
    public string? FillColor { get; init; }
    public CellFont? Font { get; init; }
    public CellBorders? Borders { get; init; }
    public string? FormatCode { get; init; }

    public static CellStyle Create(string? fillColor = null, CellFont? font = null, CellBorders? borders = null, string? formatCode = null)
    {
        return new()
        {
            FillColor = fillColor,
            Font = font,
            Borders = borders,
            FormatCode = formatCode
        };
    }
}
