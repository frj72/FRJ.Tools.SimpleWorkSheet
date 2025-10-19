using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellStyleBuilderTests
{
    [Fact]
    public void CellStyleBuilder_Create_ReturnsBuilder()
    {
        var builder = CellStyleBuilder.Create();

        Assert.NotNull(builder);
    }

    [Fact]
    public void CellStyleBuilder_WithFillColor_SetsFillColor()
    {
        var color = "FFFFFF";

        var style = CellStyleBuilder.Create()
            .WithFillColor(color)
            .Build();

        Assert.Equal(color, style.FillColor);
    }

    [Fact]
    public void CellStyleBuilder_WithFillColor_InvalidColor_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFillColor("invalidColor"));
    }

    [Fact]
    public void CellStyleBuilder_WithFont_SetsFont()
    {
        var font = CellFont.Create(14, "Arial", "000000", true, false, false, false);

        var style = CellStyleBuilder.Create()
            .WithFont(font)
            .Build();

        Assert.Equal(font, style.Font);
    }

    [Fact]
    public void CellStyleBuilder_WithFont_InvalidFontColor_ThrowsException()
    {
        var font = CellFont.Create(14, "Arial", "invalidColor", false, false, false, false);
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithFont(font));
    }

    [Fact]
    public void CellStyleBuilder_WithFontAction_ConfiguresFont()
    {
        var style = CellStyleBuilder.Create()
            .WithFont(font => font
                .WithSize(16)
                .WithName("Calibri")
                .Bold()
                .Italic())
            .Build();

        Assert.NotNull(style.Font);
        Assert.Equal(16, style.Font.Size);
        Assert.Equal("Calibri", style.Font.Name);
        Assert.True(style.Font.Bold);
        Assert.True(style.Font.Italic);
    }

    [Fact]
    public void CellStyleBuilder_WithFontAction_ModifiesExistingFont()
    {
        var existingStyle = CellStyle.Create(null, CellFont.Create(12, "Arial", "000000"), null, null);

        var style = CellStyleBuilder.FromStyle(existingStyle)
            .WithFont(font => font.WithSize(16).Bold())
            .Build();

        Assert.Equal(16, style.Font?.Size);
        Assert.Equal("Arial", style.Font?.Name);
        Assert.True(style.Font?.Bold);
    }

    [Fact]
    public void CellStyleBuilder_WithBorders_SetsBorders()
    {
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));

        var style = CellStyleBuilder.Create()
            .WithBorders(borders)
            .Build();

        Assert.Equal(borders, style.Borders);
    }

    [Fact]
    public void CellStyleBuilder_WithBorders_InvalidBorderColor_ThrowsException()
    {
        var borders = CellBorders.Create(
            CellBorder.Create("invalidColor", CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithBorders(borders));
    }

    [Fact]
    public void CellStyleBuilder_WithFormatCode_SetsFormatCode()
    {
        var formatCode = "0.000";

        var style = CellStyleBuilder.Create()
            .WithFormatCode(formatCode)
            .Build();

        Assert.Equal(formatCode, style.FormatCode);
    }

    [Fact]
    public void CellStyleBuilder_FluentChaining_SetsAllProperties()
    {
        var font = CellFont.Create(14, "Arial", "000000");
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));

        var style = CellStyleBuilder.Create()
            .WithFillColor("FFFFFF")
            .WithFont(font)
            .WithBorders(borders)
            .WithFormatCode("0.00")
            .Build();

        Assert.Equal("FFFFFF", style.FillColor);
        Assert.Equal(font, style.Font);
        Assert.Equal(borders, style.Borders);
        Assert.Equal("0.00", style.FormatCode);
    }

    [Fact]
    public void CellStyleBuilder_FromStyle_CopiesAllProperties()
    {
        var originalStyle = CellStyle.Create("FF0000", CellFont.Create(12, "Arial", "000000"), null, "0.00");

        var newStyle = CellStyleBuilder.FromStyle(originalStyle)
            .WithFillColor("00FF00")
            .Build();

        Assert.Equal("00FF00", newStyle.FillColor);
        Assert.Equal(12, newStyle.Font?.Size);
        Assert.Equal("0.00", newStyle.FormatCode);
    }

    [Fact]
    public void CellStyleBuilder_DefaultBuild_CreatesEmptyStyle()
    {
        var style = CellStyleBuilder.Create().Build();

        Assert.Null(style.FillColor);
        Assert.Null(style.Font);
        Assert.Null(style.Borders);
        Assert.Null(style.FormatCode);
    }

    [Fact]
    public void CellStyleBuilder_WithNullValues_AllowsNulls()
    {
        var style = CellStyleBuilder.Create()
            .WithFillColor(null)
            .WithFont((CellFont?)null)
            .WithBorders(null)
            .WithFormatCode(null)
            .Build();

        Assert.Null(style.FillColor);
        Assert.Null(style.Font);
        Assert.Null(style.Borders);
        Assert.Null(style.FormatCode);
    }
}
