using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellFontExtensionsTests
{
    [Fact]
    public void HasValidColor_WithNullFont_ReturnsTrue()
    {
        CellFont? font = null;

        var result = font.HasValidColor();

        Assert.True(result);
    }

    [Fact]
    public void HasValidColor_WithNullColor_ReturnsTrue()
    {
        var font = new CellFont(12, "Aptos", null, false, false, false, false);

        var result = font.HasValidColor();

        Assert.True(result);
    }

    [Fact]
    public void HasValidColor_WithValidColor_ReturnsTrue()
    {
        const string validColor = "FFAA11";
        var font = new CellFont(12, "Aptos", validColor, false, false, false, false);

        var result = font.HasValidColor();

        Assert.True(result);
    }

    [Fact]
    public void HasValidColor_WithInvalidColor_ReturnsFalse()
    {
        const string invalidColor = "XYZXYZ";
        var font = new CellFont(12, "Aptos", invalidColor, false, false, false, false);

        var result = font.HasValidColor();

        Assert.False(result);
    }
}
