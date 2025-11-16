using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellRangeTests
{
    [Fact]
    public void CellRange_FromBounds_WithReversedXCoordinates_NormalizesCorrectly()
    {
        var range = CellRange.FromBounds(5, 0, 2, 0);

        Assert.Equal(2u, range.From.X);
        Assert.Equal(5u, range.To.X);
        Assert.Equal(0u, range.From.Y);
        Assert.Equal(0u, range.To.Y);
    }

    [Fact]
    public void CellRange_FromBounds_WithReversedYCoordinates_NormalizesCorrectly()
    {
        var range = CellRange.FromBounds(0, 10, 0, 5);

        Assert.Equal(0u, range.From.X);
        Assert.Equal(0u, range.To.X);
        Assert.Equal(5u, range.From.Y);
        Assert.Equal(10u, range.To.Y);
    }

    [Fact]
    public void CellRange_SingleCell_IsSingleCellReturnsTrue()
    {
        var range = CellRange.FromBounds(3, 3, 3, 3);

        Assert.True(range.IsSingleCell);
        Assert.Equal(range.From, range.To);
    }

    [Fact]
    public void CellRange_FromPositions_WithExcelLimitPositions_StoresCorrectly()
    {
        var start = new CellPosition(0, 0);
        var end = new CellPosition(16383, 1048575);
        var range = CellRange.FromPositions(start, end);

        Assert.Equal(0u, range.From.X);
        Assert.Equal(0u, range.From.Y);
        Assert.Equal(16383u, range.To.X);
        Assert.Equal(1048575u, range.To.Y);
    }

    [Fact]
    public void CellRange_Overlaps_WithOverlappingRanges_ReturnsTrue()
    {
        var range1 = CellRange.FromBounds(0, 0, 5, 5);
        var range2 = CellRange.FromBounds(3, 3, 8, 8);

        Assert.True(range1.Overlaps(range2));
        Assert.True(range2.Overlaps(range1));
    }

    [Fact]
    public void CellRange_Overlaps_WithNonOverlappingRanges_ReturnsFalse()
    {
        var range1 = CellRange.FromBounds(0, 0, 2, 2);
        var range2 = CellRange.FromBounds(5, 5, 8, 8);

        Assert.False(range1.Overlaps(range2));
        Assert.False(range2.Overlaps(range1));
    }

    [Fact]
    public void CellRange_Overlaps_WithAdjacentRanges_ReturnsFalse()
    {
        var range1 = CellRange.FromBounds(0, 0, 2, 2);
        var range2 = CellRange.FromBounds(3, 0, 5, 2);

        Assert.False(range1.Overlaps(range2));
        Assert.False(range2.Overlaps(range1));
    }

    [Fact]
    public void CellRange_EqualityComparison_SameRanges_AreEqual()
    {
        var range1 = CellRange.FromBounds(0, 0, 5, 10);
        var range2 = CellRange.FromBounds(0, 0, 5, 10);

        Assert.Equal(range1, range2);
    }

    [Fact]
    public void CellRange_EqualityComparison_DifferentRanges_AreNotEqual()
    {
        var range1 = CellRange.FromBounds(0, 0, 5, 10);
        var range2 = CellRange.FromBounds(0, 0, 5, 11);

        Assert.NotEqual(range1, range2);
    }

    [Fact]
    public void CellRange_IsSingleCell_WithMultipleCells_ReturnsFalse()
    {
        var range = CellRange.FromBounds(0, 0, 5, 5);

        Assert.False(range.IsSingleCell);
    }
}
