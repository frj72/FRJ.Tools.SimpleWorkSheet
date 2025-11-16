using FRJ.Tools.SimpleWorkSheet.Components.Charts;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartPositionTests
{
    [Fact]
    public void ChartPosition_Constructor_SetsProperties()
    {
        var position = new ChartPosition(5, 0, 12, 15);

        Assert.Equal(5u, position.FromColumn);
        Assert.Equal(0u, position.FromRow);
        Assert.Equal(12u, position.ToColumn);
        Assert.Equal(15u, position.ToRow);
    }

    [Fact]
    public void ChartPosition_FromBounds_CreatesPosition()
    {
        var position = ChartPosition.FromBounds(5, 0, 12, 15);

        Assert.Equal(5u, position.FromColumn);
        Assert.Equal(0u, position.FromRow);
        Assert.Equal(12u, position.ToColumn);
        Assert.Equal(15u, position.ToRow);
    }

    [Fact]
    public void ChartPosition_Constructor_ThrowsWhenToColumnNotGreaterThanFromColumn()
    {
        Assert.Throws<ArgumentException>(() => new ChartPosition(5, 0, 5, 15));
        Assert.Throws<ArgumentException>(() => new ChartPosition(5, 0, 3, 15));
    }

    [Fact]
    public void ChartPosition_Constructor_ThrowsWhenToRowNotGreaterThanFromRow()
    {
        Assert.Throws<ArgumentException>(() => new ChartPosition(5, 10, 12, 10));
        Assert.Throws<ArgumentException>(() => new ChartPosition(5, 10, 12, 8));
    }

    [Fact]
    public void ChartPosition_ValidPosition_DoesNotThrow()
    {
        var exception = Record.Exception(() => new ChartPosition(0, 0, 10, 10));

        Assert.Null(exception);
    }
}
