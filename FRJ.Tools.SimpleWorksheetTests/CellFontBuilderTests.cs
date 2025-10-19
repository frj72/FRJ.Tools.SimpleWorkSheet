using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellFontBuilderTests
{
    [Fact]
    public void CellFontBuilder_Create_ReturnsBuilder()
    {
        var builder = CellFontBuilder.Create();

        Assert.NotNull(builder);
    }

    [Fact]
    public void CellFontBuilder_WithSize_SetsSize()
    {
        var size = 16;

        var font = CellFontBuilder.Create()
            .WithSize(size)
            .Build();

        Assert.Equal(size, font.Size);
    }

    [Fact]
    public void CellFontBuilder_WithName_SetsName()
    {
        var name = "Calibri";

        var font = CellFontBuilder.Create()
            .WithName(name)
            .Build();

        Assert.Equal(name, font.Name);
    }

    [Fact]
    public void CellFontBuilder_WithColor_SetsColor()
    {
        var color = "FF0000";

        var font = CellFontBuilder.Create()
            .WithColor(color)
            .Build();

        Assert.Equal(color, font.Color);
    }

    [Fact]
    public void CellFontBuilder_WithColor_InvalidColor_ThrowsException()
    {
        var builder = CellFontBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithColor("invalidColor"));
    }

    [Fact]
    public void CellFontBuilder_Bold_SetsBoldTrue()
    {
        var font = CellFontBuilder.Create()
            .Bold()
            .Build();

        Assert.True(font.Bold);
    }

    [Fact]
    public void CellFontBuilder_Bold_WithFalse_SetsBoldFalse()
    {
        var font = CellFontBuilder.Create()
            .Bold(false)
            .Build();

        Assert.False(font.Bold);
    }

    [Fact]
    public void CellFontBuilder_Italic_SetsItalicTrue()
    {
        var font = CellFontBuilder.Create()
            .Italic()
            .Build();

        Assert.True(font.Italic);
    }

    [Fact]
    public void CellFontBuilder_Italic_WithFalse_SetsItalicFalse()
    {
        var font = CellFontBuilder.Create()
            .Italic(false)
            .Build();

        Assert.False(font.Italic);
    }

    [Fact]
    public void CellFontBuilder_Underline_SetsUnderlineTrue()
    {
        var font = CellFontBuilder.Create()
            .Underline()
            .Build();

        Assert.True(font.Underline);
    }

    [Fact]
    public void CellFontBuilder_Underline_WithFalse_SetsUnderlineFalse()
    {
        var font = CellFontBuilder.Create()
            .Underline(false)
            .Build();

        Assert.False(font.Underline);
    }

    [Fact]
    public void CellFontBuilder_Strike_SetsStrikeTrue()
    {
        var font = CellFontBuilder.Create()
            .Strike()
            .Build();

        Assert.True(font.Strike);
    }

    [Fact]
    public void CellFontBuilder_Strike_WithFalse_SetsStrikeFalse()
    {
        var font = CellFontBuilder.Create()
            .Strike(false)
            .Build();

        Assert.False(font.Strike);
    }

    [Fact]
    public void CellFontBuilder_FluentChaining_SetsAllProperties()
    {
        var font = CellFontBuilder.Create()
            .WithSize(16)
            .WithName("Calibri")
            .WithColor("0000FF")
            .Bold()
            .Italic()
            .Underline()
            .Strike()
            .Build();

        Assert.Equal(16, font.Size);
        Assert.Equal("Calibri", font.Name);
        Assert.Equal("0000FF", font.Color);
        Assert.True(font.Bold);
        Assert.True(font.Italic);
        Assert.True(font.Underline);
        Assert.True(font.Strike);
    }

    [Fact]
    public void CellFontBuilder_FromFont_CopiesAllProperties()
    {
        var originalFont = CellFont.Create(14, "Arial", "FF0000", true, false, true, false);

        var newFont = CellFontBuilder.FromFont(originalFont)
            .WithSize(16)
            .Build();

        Assert.Equal(16, newFont.Size);
        Assert.Equal("Arial", newFont.Name);
        Assert.Equal("FF0000", newFont.Color);
        Assert.True(newFont.Bold);
        Assert.False(newFont.Italic);
        Assert.True(newFont.Underline);
        Assert.False(newFont.Strike);
    }

    [Fact]
    public void CellFontBuilder_DefaultBuild_UsesWorkSheetDefaults()
    {
        var font = CellFontBuilder.Create().Build();

        Assert.Equal(WorkSheetDefaults.FontSize, font.Size);
        Assert.Equal(WorkSheetDefaults.FontName, font.Name);
        Assert.Equal(WorkSheetDefaults.Color, font.Color);
        Assert.False(font.Bold);
        Assert.False(font.Italic);
        Assert.False(font.Underline);
        Assert.False(font.Strike);
    }

    [Fact]
    public void CellFontBuilder_MultipleStyleFlags_AllWork()
    {
        var font = CellFontBuilder.Create()
            .Bold()
            .Italic()
            .Build();

        Assert.True(font.Bold);
        Assert.True(font.Italic);
        Assert.False(font.Underline);
        Assert.False(font.Strike);
    }

    [Fact]
    public void CellFontBuilder_ToggleFlags_WorksCorrectly()
    {
        var font = CellFontBuilder.Create()
            .Bold()
            .Bold(false)
            .Italic()
            .Build();

        Assert.False(font.Bold);
        Assert.True(font.Italic);
    }
}
