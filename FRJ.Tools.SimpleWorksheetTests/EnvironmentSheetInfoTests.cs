using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class EnvironmentSheetInfoTests
{
    [Fact]
    public void GetWidth_WithShortString_ReturnsValidWidth()
    {
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Hi");

        Assert.True(result > 0);
        Assert.True(result < 20);
    }

    [Fact]
    public void GetWidth_WithLongString_ReturnsLargerWidth()
    {
        var shortResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Hi");
        var longResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "This is a very long string that should require more width");

        Assert.True(longResult > shortResult);
        Assert.True(longResult > 10);
    }

    [Fact]
    public void GetWidth_WithEmptyString_ReturnsDefaultWidth()
    {
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "");

        Assert.Equal(65.0 / 6.0, result);
    }

    [Fact]
    public void GetWidth_WithLargerFontSize_ReturnsWiderWidth()
    {
        var smallResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 10, "Test");
        var largeResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 20, "Test");

        Assert.True(largeResult > smallResult);
    }

    [Fact]
    public void GetWidth_WithDifferentFontFamilies_ReturnsValidWidths()
    {
        var aptosResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Test String");
        var arialResult = EnvironmentSheetInfo.GetWidth("Arial", 12, "Test String");

        Assert.True(aptosResult > 0);
        Assert.True(arialResult > 0);
    }

    [Fact]
    public void GetWidth_WithBoldEnabled_ReturnsWiderWidth()
    {
        var normalResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Bold Test", bold: false);
        var boldResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Bold Test", bold: true);

        Assert.True(boldResult >= normalResult);
    }

    [Fact]
    public void GetWidth_WithItalicEnabled_ReturnsValidWidth()
    {
        var normalResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Italic Test", italic: false);
        var italicResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Italic Test", italic: true);

        Assert.True(italicResult > 0);
        Assert.True(Math.Abs(normalResult - italicResult) < normalResult * 0.5);
    }

    [Fact]
    public void GetWidth_WithBoldAndItalic_ReturnsValidWidth()
    {
        var normalResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Combined Test", bold: false, italic: false);
        var combinedResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Combined Test", bold: true, italic: true);

        Assert.True(combinedResult >= normalResult);
    }

    [Fact]
    public void GetWidth_WithUnicodeCharacters_HandlesCorrectly()
    {
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Hello 世界 🌍");

        Assert.True(result > 0);
        Assert.True(result < 100);
    }

    [Fact]
    public void GetWidth_WithNumbers_ReturnsValidWidth()
    {
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "1234567890");

        Assert.True(result > 0);
        Assert.True(result < 50);
    }

    [Fact]
    public void GetWidth_MatchesEstimateMaxWidth_ForSameParameters()
    {
        var fontName = "Aptos Narrow";
        var fontSize = 12;
        var text = "Test String";
        var bold = true;
        var italic = false;

        var getWidthResult = EnvironmentSheetInfo.GetWidth(fontName, fontSize, text, bold, italic);

        var font = CellFont.Create(fontSize, fontName, Colors.Black, bold, italic);
        var style = CellStyle.Create(Colors.White, font, WorkSheetDefaults.CellBorders);
        var cell = new Cell(text, style, null);
        var estimateMaxWidthResult = new[] { cell }.EstimateMaxWidth();

        Assert.Equal(estimateMaxWidthResult, getWidthResult);
    }

    [Fact]
    public void GetWidth_WithVeryLongText_HandlesCorrectly()
    {
        var longText = new string('A', 1000);
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, longText);

        Assert.True(result > 100);
    }

    [Fact]
    public void GetWidth_WithSpecialCharacters_HandlesCorrectly()
    {
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Test!@#$%^&*()_+-=[]{}|;':\",./<>?");

        Assert.True(result > 0);
        Assert.True(result < 200);
    }

    [Fact]
    public void GetWidth_WithSmallFontSize_ReturnsSmallWidth()
    {
        var largeResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Test");
        var smallResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 6, "Test");

        Assert.True(smallResult < largeResult);
        Assert.True(smallResult > 0);
    }

    [Fact]
    public void GetWidth_WithVeryLargeFontSize_ReturnsLargeWidth()
    {
        var normalResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Test");
        var largeResult = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 72, "Test");

        Assert.True(largeResult > normalResult * 5);
    }

    [Fact]
    public void GetWidth_WithWhitespaceOnly_ReturnsValidWidth()
    {
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "     ");

        Assert.True(result > 0);
        Assert.True(result < 50);
    }

    [Fact]
    public void GetWidth_WithMixedContent_ReturnsValidWidth()
    {
        var result = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Test 123 ABC !@#");

        Assert.True(result > 0);
        Assert.True(result < 100);
    }

    [Fact]
    public void GetWidth_DefaultParameters_UsesCorrectDefaults()
    {
        var withExplicitDefaults = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Test", bold: false, italic: false);
        var withImplicitDefaults = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, "Test");

        Assert.Equal(withExplicitDefaults, withImplicitDefaults);
    }
}
