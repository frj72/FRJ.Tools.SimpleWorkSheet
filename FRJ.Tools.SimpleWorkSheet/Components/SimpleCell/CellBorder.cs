using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public record CellBorder(string? Color, CellBorderStyle Style)
{
    public static CellBorder Create(string? color, CellBorderStyle style) => new(color ?? WorkSheetDefaults.Color, style);
};
