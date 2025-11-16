using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ColorsTests
{
    [Fact]
    public void Colors_Black_ReturnsCorrectValue()
    {
        Assert.Equal("FF000000", Colors.Black);
    }

    [Fact]
    public void Colors_White_ReturnsCorrectValue()
    {
        Assert.Equal("FFFFFFFF", Colors.White);
    }

    [Fact]
    public void Colors_Red_ReturnsCorrectValue()
    {
        Assert.Equal("FFFF0000", Colors.Red);
    }

    [Fact]
    public void Colors_Green_ReturnsCorrectValue()
    {
        Assert.Equal("FF00FF00", Colors.Green);
    }

    [Fact]
    public void Colors_Blue_ReturnsCorrectValue()
    {
        Assert.Equal("FF0000FF", Colors.Blue);
    }

    [Fact]
    public void Colors_Yellow_ReturnsCorrectValue()
    {
        Assert.Equal("FFFFFF00", Colors.Yellow);
    }

    [Fact]
    public void Colors_Cyan_ReturnsCorrectValue()
    {
        Assert.Equal("FF00FFFF", Colors.Cyan);
    }

    [Fact]
    public void Colors_Magenta_ReturnsCorrectValue()
    {
        Assert.Equal("FFFF00FF", Colors.Magenta);
    }

    [Fact]
    public void ColorValidation_ValidSixCharHex_DoesNotThrow()
    {
        var builder = CellStyleBuilder.Create();

        var exception = Record.Exception(() => builder.WithFillColor("FF0000"));

        Assert.Null(exception);
    }

    [Fact]
    public void ColorValidation_ValidEightCharArgb_DoesNotThrow()
    {
        var builder = CellStyleBuilder.Create();

        var exception = Record.Exception(() => builder.WithFillColor("FFFF0000"));

        Assert.Null(exception);
    }

    [Fact]
    public void ColorValidation_LowercaseHex_DoesNotThrow()
    {
        var builder = CellStyleBuilder.Create();

        var exception = Record.Exception(() => builder.WithFillColor("ff0000"));

        Assert.Null(exception);
    }

    [Fact]
    public void ColorValidation_InvalidHexCharacters_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFillColor("GGGGGG"));
    }

    [Fact]
    public void ColorValidation_TooShortHex_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFillColor("FFF"));
    }

    [Fact]
    public void ColorValidation_TooLongHex_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFillColor("FFFFFFFFF"));
    }

    [Fact]
    public void ColorValidation_EmptyString_DoesNotThrow()
    {
        var builder = CellStyleBuilder.Create();

        var exception = Record.Exception(() => builder.WithFillColor(""));

        Assert.Null(exception);
    }

    [Fact]
    public void ColorValidation_NullString_DoesNotThrow()
    {
        var builder = CellStyleBuilder.Create();

        var exception = Record.Exception(() => builder.WithFillColor(null));

        Assert.Null(exception);
    }

    [Fact]
    public void ColorValidation_WithHashPrefix_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFillColor("#FF0000"));
    }

    [Fact]
    public void ColorValidation_WithSpaces_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFillColor(" FF0000 "));
    }

    [Fact]
    public void CellBuilder_WithInvalidColor_ThrowsArgumentException()
    {
        var builder = CellBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithColor("GGGGGG"));
    }

    [Fact]
    public void CellStyleBuilder_WithInvalidFillColor_ThrowsArgumentException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFillColor("GGGGGG"));
    }

}
