using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartSeriesColorTests
{
    [Fact]
    public void ChartSeries_Constructor_WithValidRgbColor_SetsColor()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);
        var series = new ChartSeries("Test Series", range, "FF0000");

        Assert.Equal("FF0000", series.Color);
    }

    [Fact]
    public void ChartSeries_Constructor_WithValidArgbColor_SetsColor()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);
        var series = new ChartSeries("Test Series", range, "FFFF0000");

        Assert.Equal("FFFF0000", series.Color);
    }

    [Fact]
    public void ChartSeries_Constructor_WithNullColor_SetsNullColor()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);
        var series = new ChartSeries("Test Series", range);

        Assert.Null(series.Color);
    }

    [Fact]
    public void ChartSeries_Constructor_WithInvalidColor_ThrowsArgumentException()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);

        Assert.Throws<ArgumentException>(() => new ChartSeries("Test Series", range, "INVALID"));
        Assert.Throws<ArgumentException>(() => new ChartSeries("Test Series", range, "123"));
    }
}
