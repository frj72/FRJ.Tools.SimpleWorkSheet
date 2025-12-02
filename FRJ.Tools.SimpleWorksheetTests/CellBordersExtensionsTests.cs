using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellBordersExtensionsTests
{
    [Fact]
    public void HasValidColor_WithNullBorder_ReturnsTrue()
    {
        CellBorder? border = null;

        var result = border.HasValidColor();

        Assert.True(result);
    }

    [Fact]
    public void HasValidColor_WithInvalidHexColor_ReturnsFalse()
    {
        const string invalidColor = "GGGGGG";
        var border = new CellBorder(invalidColor, CellBorderStyle.Thin);

        var result = border.HasValidColor();

        Assert.False(result);
    }

    [Fact]
    public void HasValidColor_WithValidHexColor_ReturnsTrue()
    {
        const string validColor = "FFA1B2";
        var border = new CellBorder(validColor, CellBorderStyle.Medium);

        var result = border.HasValidColor();

        Assert.True(result);
    }

    [Fact]
    public void HasValidColors_WithNullCellBorders_ReturnsTrue()
    {
        CellBorders? borders = null;

        var result = borders.HasValidColors();

        Assert.True(result);
    }

    [Fact]
    public void HasValidColors_WithInvalidSide_ReturnsFalse()
    {
        const string validColor = "FFA1B2";
        const string invalidColor = "12345";
        var borders = new CellBorders(
            new(validColor, CellBorderStyle.Thin),
            new(validColor, CellBorderStyle.Thin),
            new(validColor, CellBorderStyle.Thin),
            new(invalidColor, CellBorderStyle.Thin));

        var result = borders.HasValidColors();

        Assert.False(result);
    }

    [Fact]
    public void GetAllColors_WithDuplicateEntries_ReturnsUniqueColors()
    {
        const string red = "FF0000";
        const string blue = "0000FF";
        var borders = new CellBorders(
            new(red, CellBorderStyle.Thin),
            new(red, CellBorderStyle.Medium),
            new(blue, CellBorderStyle.Thick),
            null);

        var colors = borders.GetAllColors().ToList();

        Assert.Equal(2, colors.Count);
        Assert.Contains(red, colors);
        Assert.Contains(blue, colors);
    }
}
