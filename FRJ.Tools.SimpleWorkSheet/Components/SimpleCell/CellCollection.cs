namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public class CellCollection
{
    public Dictionary<CellPosition, Cell> Cells { get; } = new();

    public void SetCell(CellPosition position, Cell cell)
    {
        Cells[position] = cell;
    }
}
