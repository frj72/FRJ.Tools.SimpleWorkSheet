namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellFontExtensions
{ 
    public static bool HasValidColor(this CellFont? font) => font is null || font.Color.IsValidColor();
}
