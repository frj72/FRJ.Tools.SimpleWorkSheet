using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellColorExtensionsTests
{
    [Fact]
    public void IsValidColor_WithNullString_ReturnsTrue()
    {
        string? color = null;

        var result = color.IsValidColor();

        Assert.True(result);
    }

    [Fact]
    public void IsValidColor_WithSixCharacterHex_ReturnsTrue()
    {
        const string color = "A1B2C3";

        var result = color.IsValidColor();

        Assert.True(result);
    }

    [Fact]
    public void IsValidColor_WithEightCharacterHex_ReturnsTrue()
    {
        const string color = "FFA1B2C3";

        var result = color.IsValidColor();

        Assert.True(result);
    }

    [Fact]
    public void IsValidColor_WithInvalidLength_ReturnsFalse()
    {
        const string color = "12345";

        var result = color.IsValidColor();

        Assert.False(result);
    }

    [Fact]
    public void IsValidColor_WithInvalidCharacters_ReturnsFalse()
    {
        const string color = "GHIJKL";

        var result = color.IsValidColor();

        Assert.False(result);
    }

    [Fact]
    public void ToArgbColor_WithNullString_ReturnsBlack()
    {
        string? color = null;

        var result = color.ToArgbColor();

        Assert.Equal("FF000000", result);
    }

    [Fact]
    public void ToArgbColor_WithSixCharacterHex_AddsOpaquePrefix()
    {
        const string color = "A1B2C3";

        var result = color.ToArgbColor();

        Assert.Equal("FFA1B2C3", result);
    }

    [Fact]
    public void ToArgbColor_WithEightCharacterHex_ReturnsSameValue()
    {
        const string color = "80FFA1B2";

        var result = color.ToArgbColor();

        Assert.Equal(color, result);
    }

    [Fact]
    public void ToArgbColor_WithInvalidLength_ThrowsArgumentException()
    {
        const string color = "ABC";

        Assert.Throws<ArgumentException>(() => color.ToArgbColor());
    }
}
