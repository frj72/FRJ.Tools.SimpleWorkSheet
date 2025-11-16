using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellPositionTests
{
    [Fact]
    public void CellPosition_WithZeroZero_IsFirstCell()
    {
        var position = new CellPosition(0, 0);

        Assert.Equal(0u, position.X);
        Assert.Equal(0u, position.Y);
    }

    [Fact]
    public void CellPosition_WithLastValidColumn_StoresCorrectly()
    {
        var position = new CellPosition(16383, 0);

        Assert.Equal(16383u, position.X);
        Assert.Equal(0u, position.Y);
    }

    [Fact]
    public void CellPosition_WithLastValidRow_StoresCorrectly()
    {
        var position = new CellPosition(0, 1048575);

        Assert.Equal(0u, position.X);
        Assert.Equal(1048575u, position.Y);
    }

    [Fact]
    public void CellPosition_WithMaxUIntColumn_StoresValue()
    {
        var position = new CellPosition(uint.MaxValue, 0);

        Assert.Equal(uint.MaxValue, position.X);
        Assert.Equal(0u, position.Y);
    }

    [Fact]
    public void CellPosition_WithMaxUIntRow_StoresValue()
    {
        var position = new CellPosition(0, uint.MaxValue);

        Assert.Equal(0u, position.X);
        Assert.Equal(uint.MaxValue, position.Y);
    }

    [Fact]
    public void CellPosition_EqualityComparison_SameValues_AreEqual()
    {
        var position1 = new CellPosition(5, 10);
        var position2 = new CellPosition(5, 10);

        Assert.Equal(position1, position2);
        Assert.True(position1 == position2);
        Assert.False(position1 != position2);
    }

    [Fact]
    public void CellPosition_EqualityComparison_DifferentValues_AreNotEqual()
    {
        var position1 = new CellPosition(5, 10);
        var position2 = new CellPosition(5, 11);

        Assert.NotEqual(position1, position2);
        Assert.False(position1 == position2);
        Assert.True(position1 != position2);
    }

    [Fact]
    public void CellPosition_GetHashCode_SameValues_ProduceSameHash()
    {
        var position1 = new CellPosition(5, 10);
        var position2 = new CellPosition(5, 10);

        Assert.Equal(position1.GetHashCode(), position2.GetHashCode());
    }
}
