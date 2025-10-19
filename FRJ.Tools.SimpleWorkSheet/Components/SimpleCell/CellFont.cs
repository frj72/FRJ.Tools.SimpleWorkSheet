using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public record CellFont(int? Size, string? Name, string? Color, bool Bold, bool Italic, bool Underline, bool Strike)
{
    public static CellFont Create(int? size = null, string? name = null, string? color = null, 
        bool bold = false, bool italic = false, bool underline = false, bool strike = false) 
        => new(size ?? WorkSheetDefaults.FontSize, name ?? WorkSheetDefaults.FontName, 
            color ?? WorkSheetDefaults.Color, bold, italic, underline, strike);

    public static CellFont Create(string name) => 
        Create(null, name, null, false, false, false, false);
};
