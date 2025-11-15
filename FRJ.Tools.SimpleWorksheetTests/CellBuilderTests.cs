using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellBuilderTests
{
    [Fact]
    public void CellBuilder_Create_ReturnsBuilder()
    {
        var builder = CellBuilder.Create();

        Assert.NotNull(builder);
    }

    [Fact]
    public void CellBuilder_FromValue_SetsValue()
    {
        var value = new CellValue("TestValue");

        var cell = CellBuilder.FromValue(value).Build();

        Assert.Equal(value, cell.Value);
    }

    [Fact]
    public void CellBuilder_WithValue_SetsValue()
    {
        var value = new CellValue("TestValue");

        var cell = CellBuilder.Create()
            .WithValue(value)
            .Build();

        Assert.Equal(value, cell.Value);
    }

    [Fact]
    public void CellBuilder_WithColor_SetsFillColor()
    {
        const string color = "FF0000";

        var cell = CellBuilder.FromValue("Test")
            .WithColor(color)
            .Build();

        Assert.Equal(color, cell.Style.FillColor);
        Assert.Equal(color, cell.Color);
    }

    [Fact]
    public void CellBuilder_WithColor_InvalidColor_ThrowsException()
    {
        var builder = CellBuilder.FromValue("Test");

        Assert.Throws<ArgumentException>(() => builder.WithColor("invalidColor"));
    }

    [Fact]
    public void CellBuilder_WithFont_SetsFont()
    {
        var font = CellFont.Create(14, "Arial", "000000", true);

        var cell = CellBuilder.FromValue("Test")
            .WithFont(font)
            .Build();

        Assert.Equal(font, cell.Style.Font);
        Assert.Equal(font, cell.Font);
    }

    [Fact]
    public void CellBuilder_WithFont_InvalidFontColor_ThrowsException()
    {
        var font = CellFont.Create(14, "Arial", "invalidColor");
        var builder = CellBuilder.FromValue("Test");

        Assert.Throws<ArgumentException>(() => builder.WithFont(font));
    }

    [Fact]
    public void CellBuilder_WithFontAction_ConfiguresFont()
    {
        var cell = CellBuilder.FromValue("Test")
            .WithFont(font => font
                .WithSize(16)
                .WithName("Calibri")
                .Bold()
                .Italic())
            .Build();

        Assert.NotNull(cell.Font);
        Assert.Equal(16, cell.Font.Size);
        Assert.Equal("Calibri", cell.Font.Name);
        Assert.True(cell.Font.Bold);
        Assert.True(cell.Font.Italic);
    }

    [Fact]
    public void CellBuilder_WithBorders_SetsBorders()
    {
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));

        var cell = CellBuilder.FromValue("Test")
            .WithBorders(borders)
            .Build();

        Assert.Equal(borders, cell.Style.Borders);
        Assert.Equal(borders, cell.Borders);
    }

    [Fact]
    public void CellBuilder_WithFormatCode_SetsFormatCode()
    {
        const string formatCode = "0.000";

        var cell = CellBuilder.FromValue(123.456m)
            .WithFormatCode(formatCode)
            .Build();

        Assert.Equal(formatCode, cell.Style.FormatCode);
        Assert.Equal(formatCode, cell.FormatCode);
    }

    [Fact]
    public void CellBuilder_WithStyleAction_ConfiguresStyle()
    {
        var cell = CellBuilder.FromValue("Test")
            .WithStyle(style => style
                .WithFillColor("FFFFFF")
                .WithFont(font => font.WithSize(14).Bold())
                .WithFormatCode("0.00"))
            .Build();

        Assert.Equal("FFFFFF", cell.Style.FillColor);
        Assert.Equal(14, cell.Style.Font?.Size);
        Assert.True(cell.Style.Font?.Bold);
        Assert.Equal("0.00", cell.Style.FormatCode);
    }

    [Fact]
    public void CellBuilder_WithStyle_SetsStyle()
    {
        var style = CellStyle.Create("FFFFFF", CellFont.Create(12, "Arial", "000000"), null, "0.00");

        var cell = CellBuilder.FromValue("Test")
            .WithStyle(style)
            .Build();

        Assert.Equal(style, cell.Style);
    }

    [Fact]
    public void CellBuilder_WithMetadataAction_ConfiguresMetadata()
    {
        var importedAt = DateTime.UtcNow;

        var cell = CellBuilder.FromValue("Test")
            .WithMetadata(meta => meta
                .WithSource("csv")
                .WithImportedAt(importedAt)
                .WithOriginalValue("raw_value"))
            .Build();

        Assert.NotNull(cell.Metadata);
        Assert.Equal("csv", cell.Metadata.Source);
        Assert.Equal(importedAt, cell.Metadata.ImportedAt);
        Assert.Equal("raw_value", cell.Metadata.OriginalValue);
    }

    [Fact]
    public void CellBuilder_WithMetadata_SetsMetadata()
    {
        var metadata = CellMetadata.Create("json", DateTime.UtcNow, "original");

        var cell = CellBuilder.FromValue("Test")
            .WithMetadata(metadata)
            .Build();

        Assert.Equal(metadata, cell.Metadata);
    }

    [Fact]
    public void CellBuilder_FromSource_SetsMetadataSource()
    {
        var cell = CellBuilder.FromValue("Test")
            .FromSource("csv")
            .Build();

        Assert.NotNull(cell.Metadata);
        Assert.Equal("csv", cell.Metadata.Source);
    }

    [Fact]
    public void CellBuilder_FluentChaining_WorksCorrectly()
    {
        var cell = CellBuilder.FromValue("Test")
            .WithColor("FFFFFF")
            .WithFont(font => font.WithSize(14).Bold())
            .WithFormatCode("0.00")
            .FromSource("manual")
            .Build();

        Assert.Equal("Test", cell.Value.AsString());
        Assert.Equal("FFFFFF", cell.Color);
        Assert.Equal(14, cell.Font?.Size);
        Assert.True(cell.Font?.Bold);
        Assert.Equal("0.00", cell.FormatCode);
        Assert.Equal("manual", cell.Metadata?.Source);
    }

    [Fact]
    public void CellBuilder_FromCell_CopiesAllProperties()
    {
        var originalCell = CellBuilder.FromValue("Original")
            .WithColor("FF0000")
            .WithFont(CellFont.Create(12, "Arial", "000000"))
            .WithMetadata(CellMetadata.Create("csv", DateTime.UtcNow, "raw"))
            .Build();

        var newCell = CellBuilder.FromCell(originalCell)
            .WithValue("Modified")
            .Build();

        Assert.Equal("Modified", newCell.Value.AsString());
        Assert.Equal("FF0000", newCell.Color);
        Assert.Equal(12, newCell.Font?.Size);
        Assert.NotNull(newCell.Metadata);
    }

    [Fact]
    public void CellBuilder_DefaultBuild_UsesDefaults()
    {
        var cell = CellBuilder.Create().Build();

        Assert.Equal(string.Empty, cell.Value.AsString());
        Assert.Equal(WorkSheetDefaults.DefaultCellStyle, cell.Style);
    }
}
