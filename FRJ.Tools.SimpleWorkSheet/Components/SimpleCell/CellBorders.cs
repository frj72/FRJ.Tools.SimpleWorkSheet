namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public record CellBorders(CellBorder? Left, CellBorder? Right, CellBorder? Top, CellBorder? Bottom)
{
    public static CellBorders Create(CellBorder? left, CellBorder? right, CellBorder? top, CellBorder? bottom) => new(left ?? CellBorder.Create(null, CellBorderStyle.None), right ?? CellBorder.Create(null, CellBorderStyle.None), top ?? CellBorder.Create(null, CellBorderStyle.None), bottom ?? CellBorder.Create(null, CellBorderStyle.None));
}
