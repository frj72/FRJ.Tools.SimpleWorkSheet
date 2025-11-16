using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public readonly record struct CellRange(CellPosition From, CellPosition To)
{
    public bool IsSingleCell => From.X == To.X && From.Y == To.Y;

    public static CellRange FromBounds(uint fromX, uint fromY, uint toX, uint toY)
    {
        var minX = Math.Min(fromX, toX);
        var minY = Math.Min(fromY, toY);
        var maxX = Math.Max(fromX, toX);
        var maxY = Math.Max(fromY, toY);
        return new(new(minX, minY), new(maxX, maxY));
    }

    public static CellRange FromPositions(CellPosition start, CellPosition end) =>
        FromBounds(start.X, start.Y, end.X, end.Y);

    public bool Overlaps(CellRange other)
    {
        var separatedHorizontally = other.To.X < From.X || other.From.X > To.X;
        var separatedVertically = other.To.Y < From.Y || other.From.Y > To.Y;
        return !separatedHorizontally && !separatedVertically;
    }
}
