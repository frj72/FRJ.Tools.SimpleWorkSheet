using FRJ.Tools.SimpleWorkSheet.Components.Charts;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartSizeTests
{
    [Fact]
    public void ChartSize_Constructor_SetsProperties()
    {
        var size = new ChartSize(6000000, 4000000);

        Assert.Equal(6000000, size.WidthEmus);
        Assert.Equal(4000000, size.HeightEmus);
    }

    [Fact]
    public void ChartSize_FromEMUs_CreatesSize()
    {
        var size = ChartSize.FromEmus(5000000, 3000000);

        Assert.Equal(5000000, size.WidthEmus);
        Assert.Equal(3000000, size.HeightEmus);
    }

    [Fact]
    public void ChartSize_FromInches_ConvertsCorrectly()
    {
        var size = ChartSize.FromInches(1.0, 2.0);

        Assert.Equal(914400, size.WidthEmus);
        Assert.Equal(1828800, size.HeightEmus);
    }

    [Fact]
    public void ChartSize_Default_HasReasonableSize()
    {
        var size = ChartSize.Default;

        Assert.True(size.WidthEmus > 0);
        Assert.True(size.HeightEmus > 0);
    }

    [Fact]
    public void ChartSize_Constructor_ThrowsOnZeroOrNegativeWidth()
    {
        Assert.Throws<ArgumentException>(() => new ChartSize(0, 4000000));
        Assert.Throws<ArgumentException>(() => new ChartSize(-1000, 4000000));
    }

    [Fact]
    public void ChartSize_Constructor_ThrowsOnZeroOrNegativeHeight()
    {
        Assert.Throws<ArgumentException>(() => new ChartSize(6000000, 0));
        Assert.Throws<ArgumentException>(() => new ChartSize(6000000, -1000));
    }

    [Fact]
    public void ChartSize_FromInches_ThrowsOnZeroOrNegativeWidth()
    {
        Assert.Throws<ArgumentException>(() => ChartSize.FromInches(0, 4));
        Assert.Throws<ArgumentException>(() => ChartSize.FromInches(-1, 4));
    }

    [Fact]
    public void ChartSize_FromInches_ThrowsOnZeroOrNegativeHeight()
    {
        Assert.Throws<ArgumentException>(() => ChartSize.FromInches(6, 0));
        Assert.Throws<ArgumentException>(() => ChartSize.FromInches(6, -1));
    }
}
