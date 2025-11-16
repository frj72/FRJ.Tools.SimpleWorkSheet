using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ChartSeriesTests
{
    [Fact]
    public void ChartSeries_Constructor_SetsProperties()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);
        var series = new ChartSeries("Test Series", range);

        Assert.Equal("Test Series", series.Name);
        Assert.Equal(range, series.DataRange);
    }

    [Fact]
    public void ChartSeries_Constructor_ThrowsOnNullName()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);
        string? nullName = null;

        Assert.Throws<ArgumentException>(() => new ChartSeries("", range));
        Assert.Throws<ArgumentException>(() => new ChartSeries(nullName, range));
    }

    [Fact]
    public void ChartDataRange_ValidateDataRange_ThrowsOnSingleCell()
    {
        var singleCell = CellRange.FromBounds(0, 0, 0, 0);

        Assert.Throws<ArgumentException>(() => ChartDataRange.ValidateDataRange(singleCell));
    }

    [Fact]
    public void ChartDataRange_ValidateDataRange_AcceptsMultipleCells()
    {
        var range = CellRange.FromBounds(0, 0, 0, 5);

        var exception = Record.Exception(() => ChartDataRange.ValidateDataRange(range));

        Assert.Null(exception);
    }

    [Fact]
    public void ChartDataRange_ToRangeReference_GeneratesCorrectFormat()
    {
        var range = CellRange.FromBounds(0, 0, 2, 5);

        var reference = ChartDataRange.ToRangeReference(range, "Sheet1");

        Assert.Equal("'Sheet1'!$A$1:$C$6", reference);
    }

    [Fact]
    public void ChartDataRange_ToRangeReference_HandlesLargeColumns()
    {
        var range = CellRange.FromBounds(25, 0, 27, 5);

        var reference = ChartDataRange.ToRangeReference(range, "Sheet1");

        Assert.Equal("'Sheet1'!$Z$1:$AB$6", reference);
    }
}
