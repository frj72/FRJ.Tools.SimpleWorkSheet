namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellBordersExtensions
{
   
    public static bool HasValidColor(this CellBorder? border) => border?.Color is null || border.Color.IsValidColor();

    public static bool HasValidColors(this CellBorders? cellBorders) =>
        cellBorders is null ||(
            cellBorders.Left.HasValidColor() &&
            cellBorders.Right.HasValidColor() &&
            cellBorders.Top.HasValidColor() &&
            cellBorders.Bottom.HasValidColor());

    public static IEnumerable<string> GetAllColors(this CellBorders cellBorders)
    {
        List<string> colors = [];
        if (cellBorders.Left?.Color is not null) colors.Add(cellBorders.Left.Color);
        if (cellBorders.Right?.Color is not null) colors.Add(cellBorders.Right.Color);
        if (cellBorders.Top?.Color is not null) colors.Add(cellBorders.Top.Color);
        if (cellBorders.Bottom?.Color is not null) colors.Add(cellBorders.Bottom.Color);
        return colors.Distinct();
    }
}
