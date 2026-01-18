using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

[Collection("TypefaceCache")]
public class CellExtensionsEstimateMaxWidthTests : IDisposable
{
    public void Dispose()
    {
        TypefaceCache.ClearCache();
        GC.SuppressFinalize(this);
    }
    [Fact]
    public void EstimateMaxWidth_WithEmptyCollection_ReturnsDefaultWidth()
    {
        var cells = Array.Empty<Cell>();

        var result = cells.EstimateMaxWidth();

        Assert.Equal(65.0 / 6.0, result);
    }

    [Fact]
    public void EstimateMaxWidth_WithSingleEmptyCell_ReturnsDefaultWidth()
    {
        var cell = new Cell("", WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.Equal(65.0 / 6.0, result);
    }

    [Fact]
    public void EstimateMaxWidth_WithShortString_ReturnsSmallWidth()
    {
        var cell = new Cell("Hi", WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
        Assert.True(result < 20);
    }

    [Fact]
    public void EstimateMaxWidth_WithLongString_ReturnsLargeWidth()
    {
        var cell = new Cell("This is a very long string that should require more width", WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 10);
    }

    [Fact]
    public void EstimateMaxWidth_WithDecimalValue_IncludesDecimalPlaces()
    {
        var cell = new Cell(123.456m, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithLongValue_IncludesSpace()
    {
        var cell = new Cell(123456789L, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithDateTimeValue_UsesFullFormat()
    {
        var dateTime = new DateTime(2024, 12, 2, 15, 30, 45);
        var cell = new Cell(dateTime, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 15);
    }

    [Fact]
    public void EstimateMaxWidth_WithLargerFontSize_ReturnsWiderWidth()
    {
        var smallFontCell = new Cell("Test", new CellStyle { Font = CellFont.Create(10) }, null);
        var largeFontCell = new Cell("Test", new CellStyle { Font = CellFont.Create(20) }, null);

        var smallResult = new[] { smallFontCell }.EstimateMaxWidth();
        var largeResult = new[] { largeFontCell }.EstimateMaxWidth();

        Assert.True(largeResult >= smallResult);
    }

    [Fact]
    public void EstimateMaxWidth_WithBoldFont_ReturnsValidWidth()
    {
        var normalCell = new Cell("Test", WorkSheetDefaults.DefaultCellStyle, null);
        var boldCell = new Cell("Test", new CellStyle { Font = CellFont.Create(bold: true) }, null);
        var cells = new[] { normalCell, boldCell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithItalicFont_ReturnsValidWidth()
    {
        var cell = new Cell("Test", new CellStyle { Font = CellFont.Create(italic: true) }, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithDifferentFontFamilies_ReturnsValidWidth()
    {
        var arialCell = new Cell("Test", new CellStyle { Font = CellFont.Create("Arial") }, null);
        var cells = new[] { arialCell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithMultipleCells_ReturnsMaxWidth()
    {
        var shortCell = new Cell("Hi", WorkSheetDefaults.DefaultCellStyle, null);
        var longCell = new Cell("This is a very long string", WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { shortCell, longCell };

        var result = cells.EstimateMaxWidth();

        var shortWidth = new[] { shortCell }.EstimateMaxWidth();
        var longWidth = new[] { longCell }.EstimateMaxWidth();

        Assert.Equal(longWidth, result);
        Assert.True(result > shortWidth);
    }

    [Fact]
    public void EstimateMaxWidth_WithFormula_UsesFormulaString()
    {
        var cell = new Cell(new CellFormula("=SUM(A1:A10)"), WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithUnicodeCharacters_HandlesCorrectly()
    {
        var cell = new Cell("Hello 世界 🌍", WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithVeryLongNumber_HandlesCorrectly()
    {
        var cell = new Cell(123456789012.123456789m, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 10);
    }

    [Fact]
    public void EstimateMaxWidth_WithZeroDecimal_HandlesCorrectly()
    {
        var cell = new Cell(0m, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithZeroLong_HandlesCorrectly()
    {
        var cell = new Cell(0L, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithNegativeNumber_HandlesCorrectly()
    {
        var cell = new Cell(-123.45m, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithMinDateTime_HandlesCorrectly()
    {
        var cell = new Cell(DateTime.MinValue, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithMaxDateTime_HandlesCorrectly()
    {
        var cell = new Cell(DateTime.MaxValue, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithDateTimeOffset_UsesDateTime()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 12, 2, 15, 30, 45, TimeSpan.Zero);
        var cell = new Cell(dateTimeOffset, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 15);
    }

    [Fact]
    public void EstimateMaxWidth_WithNullFont_UsesDefaults()
    {
        var cell = new Cell("Test", new CellStyle { Font = null }, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithNullFontSize_UsesDefaultSize()
    {
        var cell = new Cell("Test", new CellStyle { Font = CellFont.Create() }, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
    }

    [Fact]
    public void EstimateMaxWidth_WithMixedContentTypes_ReturnsMaxWidth()
    {
        var stringCell = new Cell("Short", WorkSheetDefaults.DefaultCellStyle, null);
        var numberCell = new Cell(123456789012345L, WorkSheetDefaults.DefaultCellStyle, null);
        var dateCell = new Cell(DateTime.Now, WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { stringCell, numberCell, dateCell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 10);
    }

    [Fact]
    public void EstimateMaxWidth_WithSameWidthCells_ReturnsThatWidth()
    {
        var cell1 = new Cell("Test", WorkSheetDefaults.DefaultCellStyle, null);
        var cell2 = new Cell("Test", WorkSheetDefaults.DefaultCellStyle, null);
        var cells = new[] { cell1, cell2 };

        var result = cells.EstimateMaxWidth();
        var singleResult = new[] { cell1 }.EstimateMaxWidth();

        Assert.Equal(singleResult, result);
    }

    [Fact]
    public void EstimateMaxWidth_WithVerySmallFontSize_ReturnsSmallWidth()
    {
        var cell = new Cell("Test", new CellStyle { Font = CellFont.Create(8) }, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 0);
        Assert.True(result < 10);
    }

    [Fact]
    public void EstimateMaxWidth_WithVeryLargeFontSize_ReturnsLargeWidth()
    {
        var cell = new Cell("Test", new CellStyle { Font = CellFont.Create(72) }, null);
        var cells = new[] { cell };

        var result = cells.EstimateMaxWidth();

        Assert.True(result > 20);
    }
}