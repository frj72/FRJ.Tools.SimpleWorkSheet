// ReSharper disable ConvertToExtensionBlock
namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellStyleExtensions
{
    public static bool HasValidColors(this CellStyle? style)
    {
        if (style is null)
            return true;

        return style.FillColor.IsValidColor() && 
               style.Font.HasValidColor() && 
               style.Borders.HasValidColors();
    }

    public static CellStyle WithFillColor(this CellStyle style, string? fillColor) => style with { FillColor = fillColor };

    public static CellStyle WithFont(this CellStyle style, CellFont? font) => style with { Font = font };

    public static CellStyle WithBorders(this CellStyle style, CellBorders? borders) => style with { Borders = borders };

    public static CellStyle WithFormatCode(this CellStyle style, string? formatCode) => style with { FormatCode = formatCode };
}
