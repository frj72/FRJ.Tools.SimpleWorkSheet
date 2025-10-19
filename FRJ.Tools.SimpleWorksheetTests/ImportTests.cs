using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;
// ReSharper disable StringLiteralTypo
public class ImportTests
{
    [Fact]
    public void ImportOptions_Create_WithDefaults_SetsDefaultValues()
    {
        var options = ImportOptions.Create();

        Assert.True(options.PreserveOriginalValue);
        Assert.Null(options.DefaultStyle);
        Assert.Null(options.ColumnParsers);
        Assert.Null(options.SourceIdentifier);
        Assert.True(options.TrimWhitespace);
        Assert.False(options.SkipEmptyRows);
        Assert.Null(options.CustomMetadata);
    }

    [Fact]
    public void ImportOptions_Create_WithParameters_SetsAllProperties()
    {
        var style = CellStyle.Create("FFFFFF");
        var parsers = new Dictionary<int, Func<string, CellValue>>
        {
            { 0, s => new(int.Parse(s)) }
        };
        var metadata = new Dictionary<string, object> { { "key", "value" } };

        var options = ImportOptions.Create(
            preserveOriginalValue: false,
            defaultStyle: style,
            columnParsers: parsers,
            sourceIdentifier: "csv",
            trimWhitespace: false,
            skipEmptyRows: true,
            customMetadata: metadata);

        Assert.False(options.PreserveOriginalValue);
        Assert.Equal(style, options.DefaultStyle);
        Assert.Equal(parsers, options.ColumnParsers);
        Assert.Equal("csv", options.SourceIdentifier);
        Assert.False(options.TrimWhitespace);
        Assert.True(options.SkipEmptyRows);
        Assert.Equal(metadata, options.CustomMetadata);
    }

    [Fact]
    public void ImportOptionsBuilder_Create_ReturnsBuilder()
    {
        var builder = ImportOptionsBuilder.Create();

        Assert.NotNull(builder);
    }

    [Fact]
    public void ImportOptionsBuilder_WithPreserveOriginalValue_SetsValue()
    {
        var options = ImportOptionsBuilder.Create()
            .WithPreserveOriginalValue(false)
            .Build();

        Assert.False(options.PreserveOriginalValue);
    }

    [Fact]
    public void ImportOptionsBuilder_WithDefaultStyle_SetsStyle()
    {
        var style = CellStyle.Create("FFFFFF");

        var options = ImportOptionsBuilder.Create()
            .WithDefaultStyle(style)
            .Build();

        Assert.Equal(style, options.DefaultStyle);
    }

    [Fact]
    public void ImportOptionsBuilder_WithDefaultStyleAction_ConfiguresStyle()
    {
        var options = ImportOptionsBuilder.Create()
            .WithDefaultStyle(style => style
                .WithFillColor("EFEFEF")
                .WithFont(font => font.WithSize(14).Bold()))
            .Build();

        Assert.NotNull(options.DefaultStyle);
        Assert.Equal("EFEFEF", options.DefaultStyle.FillColor);
        Assert.Equal(14, options.DefaultStyle.Font?.Size);
        Assert.True(options.DefaultStyle.Font?.Bold);
    }

    [Fact]
    public void ImportOptionsBuilder_WithColumnParser_AddsParser()
    {
        var options = ImportOptionsBuilder.Create()
            .WithColumnParser(0, s => new(int.Parse(s)))
            .WithColumnParser(1, s => new(decimal.Parse(s)))
            .Build();

        Assert.NotNull(options.ColumnParsers);
        Assert.Equal(2, options.ColumnParsers.Count);
        Assert.Equal(42, options.ColumnParsers[0]("42").AsInt());
        Assert.Equal(3.14m, options.ColumnParsers[1]("3.14").AsDecimal());
    }

    [Fact]
    public void ImportOptionsBuilder_WithSourceIdentifier_SetsSource()
    {
        var options = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("csv")
            .Build();

        Assert.Equal("csv", options.SourceIdentifier);
    }

    [Fact]
    public void ImportOptionsBuilder_WithTrimWhitespace_SetsValue()
    {
        var options = ImportOptionsBuilder.Create()
            .WithTrimWhitespace(false)
            .Build();

        Assert.False(options.TrimWhitespace);
    }

    [Fact]
    public void ImportOptionsBuilder_WithSkipEmptyRows_SetsValue()
    {
        var options = ImportOptionsBuilder.Create()
            .WithSkipEmptyRows(true)
            .Build();

        Assert.True(options.SkipEmptyRows);
    }

    [Fact]
    public void ImportOptionsBuilder_WithCustomMetadata_AddsMetadata()
    {
        var options = ImportOptionsBuilder.Create()
            .WithCustomMetadata("key1", "value1")
            .WithCustomMetadata("key2", 42)
            .Build();

        Assert.NotNull(options.CustomMetadata);
        Assert.Equal(2, options.CustomMetadata.Count);
        Assert.Equal("value1", options.CustomMetadata["key1"]);
        Assert.Equal(42, options.CustomMetadata["key2"]);
    }

    [Fact]
    public void ImportOptionsBuilder_FluentChaining_ConfiguresAllOptions()
    {
        var options = ImportOptionsBuilder.Create()
            .WithPreserveOriginalValue(false)
            .WithSourceIdentifier("csv")
            .WithTrimWhitespace(true)
            .WithSkipEmptyRows(true)
            .WithColumnParser(0, s => new(int.Parse(s)))
            .WithCustomMetadata("version", "1.0")
            .Build();

        Assert.False(options.PreserveOriginalValue);
        Assert.Equal("csv", options.SourceIdentifier);
        Assert.True(options.TrimWhitespace);
        Assert.True(options.SkipEmptyRows);
        Assert.NotNull(options.ColumnParsers);
        Assert.Single(options.ColumnParsers);
        Assert.NotNull(options.CustomMetadata);
        Assert.Single(options.CustomMetadata);
    }

    [Fact]
    public void CellImportExtensions_FromImportedValue_WithSource_SetsMetadata()
    {
        var cell = CellBuilder.FromValue("TestValue")
            .FromImportedValue("raw_value", "csv")
            .Build();

        Assert.NotNull(cell.Metadata);
        Assert.Equal("csv", cell.Metadata.Source);
        Assert.Equal("raw_value", cell.Metadata.OriginalValue);
        Assert.NotNull(cell.Metadata.ImportedAt);
    }

    [Fact]
    public void CellImportExtensions_FromImportedValue_WithOptions_AppliesAllSettings()
    {
        var options = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("json")
            .WithPreserveOriginalValue(true)
            .WithDefaultStyle(CellStyle.Create("EFEFEF"))
            .WithCustomMetadata("version", "1.0")
            .Build();

        var cell = CellBuilder.FromValue("TestValue")
            .FromImportedValue("raw_value", options)
            .Build();

        Assert.NotNull(cell.Metadata);
        Assert.Equal("json", cell.Metadata.Source);
        Assert.Equal("raw_value", cell.Metadata.OriginalValue);
        Assert.Equal("EFEFEF", cell.Style.FillColor);
        Assert.NotNull(cell.Metadata.CustomData);
        Assert.Equal("1.0", cell.Metadata.CustomData["version"]);
    }

    [Fact]
    public void CellImportExtensions_ProcessRawValue_WithTrimWhitespace_TrimsValue()
    {
        var options = ImportOptionsBuilder.Create()
            .WithTrimWhitespace(true)
            .Build();

        var processed = "  test  ".ProcessRawValue(options);

        Assert.Equal("test", processed);
    }

    [Fact]
    public void CellImportExtensions_ProcessRawValue_WithoutTrimWhitespace_PreservesValue()
    {
        var options = ImportOptionsBuilder.Create()
            .WithTrimWhitespace(false)
            .Build();

        var processed = "  test  ".ProcessRawValue(options);

        Assert.Equal("  test  ", processed);
    }

    [Fact]
    public void CellImportExtensions_ProcessRawValue_WithNullOptions_ReturnsUnchanged()
    {
        var processed = "  test  ".ProcessRawValue(null);

        Assert.Equal("  test  ", processed);
    }
}
