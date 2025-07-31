using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class WorkSheetDefaults
{
    public static string FillColor => Colors.White;
    public static string Color => Colors.Black;
    public static string FontName => "Aptos Narrow";
    public static int FontSize => 12;
    public static CellFont Font => CellFont.Create(FontSize, FontName, Color);
    public static CellBorder CellBorder => CellBorder.Create(null, CellBorderStyle.None);
    public static CellBorders CellBorders => CellBorders.Create(CellBorder, CellBorder, CellBorder, CellBorder);
    public static Cell DefaultCell => Cell.CreateEmpty();
    public static string DefaultDateFormat =>"yyyy-mm-dd hh:mm:ss";
}
