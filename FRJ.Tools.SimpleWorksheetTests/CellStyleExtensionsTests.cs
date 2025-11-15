using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellStyleExtensionsTests
{
    [Fact]
    public void HasValidColors_WithValidColors_ReturnsTrue()
    {
        var font = CellFont.Create(12, "Arial", "000000");
        var borders = CellBorders.Create(
            CellBorder.Create("FF0000", CellBorderStyle.Thin),
            CellBorder.Create("00FF00", CellBorderStyle.Thin),
            CellBorder.Create("0000FF", CellBorderStyle.Thin),
            CellBorder.Create("FFFF00", CellBorderStyle.Thin));
        var style = CellStyle.Create("FFFFFF", font, borders);

        var result = style.HasValidColors();

        Assert.True(result);
    }

    [Fact]
    public void HasValidColors_WithInvalidFillColor_ReturnsFalse()
    {
        var style = CellStyle.Create("invalidColor");

        var result = style.HasValidColors();

        Assert.False(result);
    }

    [Fact]
    public void HasValidColors_WithInvalidFontColor_ReturnsFalse()
    {
        var font = CellFont.Create(12, "Arial", "invalidColor");
        var style = CellStyle.Create("FFFFFF", font);

        var result = style.HasValidColors();

        Assert.False(result);
    }

    [Fact]
    public void HasValidColors_WithInvalidBorderColor_ReturnsFalse()
    {
        var borders = CellBorders.Create(
            CellBorder.Create("invalidColor", CellBorderStyle.Thin),
            CellBorder.Create("000000", CellBorderStyle.Thin),
            CellBorder.Create("000000", CellBorderStyle.Thin),
            CellBorder.Create("000000", CellBorderStyle.Thin));
        var style = CellStyle.Create("FFFFFF", null, borders);

        var result = style.HasValidColors();

        Assert.False(result);
    }

    [Fact]
    public void HasValidColors_WithNullStyle_ReturnsTrue()
    {
        CellStyle? style = null;
        var result = style.HasValidColors();
        Assert.True(result);
    }

    [Fact]
    public void WithFillColor_UpdatesFillColor_KeepsOtherProperties()
    {
        var font = CellFont.Create(12, "Arial", "000000");
        var originalStyle = CellStyle.Create("FFFFFF", font, null, "0.00");
        
        var updatedStyle = originalStyle.WithFillColor("FF0000");

        Assert.Equal("FF0000", updatedStyle.FillColor);
        Assert.Equal(font, updatedStyle.Font);
        Assert.Null(updatedStyle.Borders);
        Assert.Equal("0.00", updatedStyle.FormatCode);
    }

    [Fact]
    public void WithFont_UpdatesFont_KeepsOtherProperties()
    {
        var originalStyle = CellStyle.Create("FFFFFF", null, null, "0.00");
        var newFont = CellFont.Create(14, "Arial", "000000", true);
        
        var updatedStyle = originalStyle.WithFont(newFont);

        Assert.Equal("FFFFFF", updatedStyle.FillColor);
        Assert.Equal(newFont, updatedStyle.Font);
        Assert.Equal("0.00", updatedStyle.FormatCode);
    }

    [Fact]
    public void WithBorders_UpdatesBorders_KeepsOtherProperties()
    {
        var originalStyle = CellStyle.Create("FFFFFF", null, null, "0.00");
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        var updatedStyle = originalStyle.WithBorders(borders);

        Assert.Equal("FFFFFF", updatedStyle.FillColor);
        Assert.Equal(borders, updatedStyle.Borders);
        Assert.Equal("0.00", updatedStyle.FormatCode);
    }

    [Fact]
    public void WithFormatCode_UpdatesFormatCode_KeepsOtherProperties()
    {
        var font = CellFont.Create(12, "Arial", "000000");
        var originalStyle = CellStyle.Create("FFFFFF", font);
        
        var updatedStyle = originalStyle.WithFormatCode("0.000");

        Assert.Equal("FFFFFF", updatedStyle.FillColor);
        Assert.Equal(font, updatedStyle.Font);
        Assert.Equal("0.000", updatedStyle.FormatCode);
    }

    [Fact]
    public void ChainedWith_UpdatesMultipleProperties()
    {
        var originalStyle = CellStyle.Create();
        var font = CellFont.Create(14, "Arial", "000000");
        
        var updatedStyle = originalStyle
            .WithFillColor("FFFFFF")
            .WithFont(font)
            .WithFormatCode("0.00");

        Assert.Equal("FFFFFF", updatedStyle.FillColor);
        Assert.Equal(font, updatedStyle.Font);
        Assert.Equal("0.00", updatedStyle.FormatCode);
    }
}
