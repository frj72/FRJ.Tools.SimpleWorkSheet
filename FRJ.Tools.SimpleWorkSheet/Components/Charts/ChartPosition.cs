namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public record ChartPosition
{
    public uint FromColumn { get; init; }
    public uint FromRow { get; init; }
    public uint ToColumn { get; init; }
    public uint ToRow { get; init; }

    public ChartPosition(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        if (toColumn <= fromColumn)
            throw new ArgumentException("ToColumn must be greater than FromColumn", nameof(toColumn));
        if (toRow <= fromRow)
            throw new ArgumentException("ToRow must be greater than FromRow", nameof(toRow));

        FromColumn = fromColumn;
        FromRow = fromRow;
        ToColumn = toColumn;
        ToRow = toRow;
    }

    public static ChartPosition FromBounds(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        return new(fromColumn, fromRow, toColumn, toRow);
    }
}
