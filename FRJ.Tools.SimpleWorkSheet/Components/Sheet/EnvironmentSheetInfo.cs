using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class EnvironmentSheetInfo
{
    public static double GetWidth(string fontName, int fontSize, string text, bool bold = false, bool italic = false)
    {
        var font = CellFont.Create(fontSize, fontName, Colors.Black, bold, italic);
        var style = CellStyle.Create(Colors.White, font, WorkSheetDefaults.CellBorders);
        var cell = new Cell(text, style, null);
        
        return new[] { cell }.EstimateMaxWidth();
    }
}
