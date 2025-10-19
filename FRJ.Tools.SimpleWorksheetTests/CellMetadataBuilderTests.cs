using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellMetadataBuilderTests
{
    [Fact]
    public void CellMetadataBuilder_Create_ReturnsBuilder()
    {
        var builder = CellMetadataBuilder.Create();

        Assert.NotNull(builder);
    }

    [Fact]
    public void CellMetadataBuilder_WithSource_SetsSource()
    {
        var source = "csv";

        var metadata = CellMetadataBuilder.Create()
            .WithSource(source)
            .Build();

        Assert.Equal(source, metadata.Source);
    }

    [Fact]
    public void CellMetadataBuilder_WithImportedAt_SetsImportedAt()
    {
        var importedAt = DateTime.UtcNow;

        var metadata = CellMetadataBuilder.Create()
            .WithImportedAt(importedAt)
            .Build();

        Assert.Equal(importedAt, metadata.ImportedAt);
    }

    [Fact]
    public void CellMetadataBuilder_WithOriginalValue_SetsOriginalValue()
    {
        var originalValue = "raw_value";

        var metadata = CellMetadataBuilder.Create()
            .WithOriginalValue(originalValue)
            .Build();

        Assert.Equal(originalValue, metadata.OriginalValue);
    }

    [Fact]
    public void CellMetadataBuilder_WithCustomData_SetsCustomData()
    {
        var customData = new Dictionary<string, object> { { "key1", "value1" } };

        var metadata = CellMetadataBuilder.Create()
            .WithCustomData(customData)
            .Build();

        Assert.Equal(customData, metadata.CustomData);
    }

    [Fact]
    public void CellMetadataBuilder_AddCustomData_AddsToCustomData()
    {
        var metadata = CellMetadataBuilder.Create()
            .AddCustomData("key1", "value1")
            .AddCustomData("key2", 42)
            .Build();

        Assert.NotNull(metadata.CustomData);
        Assert.Equal("value1", metadata.CustomData["key1"]);
        Assert.Equal(42, metadata.CustomData["key2"]);
    }

    [Fact]
    public void CellMetadataBuilder_AddCustomData_CreatesCustomDataIfNull()
    {
        var metadata = CellMetadataBuilder.Create()
            .AddCustomData("key1", "value1")
            .Build();

        Assert.NotNull(metadata.CustomData);
        Assert.Single(metadata.CustomData);
        Assert.Equal("value1", metadata.CustomData["key1"]);
    }

    [Fact]
    public void CellMetadataBuilder_AddCustomData_UpdatesExistingKey()
    {
        var metadata = CellMetadataBuilder.Create()
            .AddCustomData("key1", "value1")
            .AddCustomData("key1", "value2")
            .Build();

        Assert.NotNull(metadata.CustomData);
        Assert.Single(metadata.CustomData);
        Assert.Equal("value2", metadata.CustomData["key1"]);
    }

    [Fact]
    public void CellMetadataBuilder_FluentChaining_SetsAllProperties()
    {
        var importedAt = DateTime.UtcNow;

        var metadata = CellMetadataBuilder.Create()
            .WithSource("csv")
            .WithImportedAt(importedAt)
            .WithOriginalValue("raw_value")
            .AddCustomData("row", 10)
            .AddCustomData("column", 5)
            .Build();

        Assert.Equal("csv", metadata.Source);
        Assert.Equal(importedAt, metadata.ImportedAt);
        Assert.Equal("raw_value", metadata.OriginalValue);
        Assert.NotNull(metadata.CustomData);
        Assert.Equal(10, metadata.CustomData["row"]);
        Assert.Equal(5, metadata.CustomData["column"]);
    }

    [Fact]
    public void CellMetadataBuilder_FromMetadata_CopiesAllProperties()
    {
        var customData = new Dictionary<string, object> { { "key1", "value1" } };
        var originalMetadata = CellMetadata.Create("csv", DateTime.UtcNow, "raw_value", customData);

        var newMetadata = CellMetadataBuilder.FromMetadata(originalMetadata)
            .WithSource("json")
            .Build();

        Assert.Equal("json", newMetadata.Source);
        Assert.Equal(originalMetadata.ImportedAt, newMetadata.ImportedAt);
        Assert.Equal("raw_value", newMetadata.OriginalValue);
        Assert.NotNull(newMetadata.CustomData);
        Assert.Equal("value1", newMetadata.CustomData["key1"]);
    }

    [Fact]
    public void CellMetadataBuilder_FromMetadata_CustomDataIsDeepCopied()
    {
        var customData = new Dictionary<string, object> { { "key1", "value1" } };
        var originalMetadata = CellMetadata.Create("csv", null, null, customData);

        var newMetadata = CellMetadataBuilder.FromMetadata(originalMetadata)
            .AddCustomData("key2", "value2")
            .Build();

        Assert.NotNull(newMetadata.CustomData);
        Assert.Equal(2, newMetadata.CustomData.Count);
        Assert.Single(originalMetadata.CustomData!);
    }

    [Fact]
    public void CellMetadataBuilder_DefaultBuild_CreatesEmptyMetadata()
    {
        var metadata = CellMetadataBuilder.Create().Build();

        Assert.Null(metadata.Source);
        Assert.Null(metadata.ImportedAt);
        Assert.Null(metadata.OriginalValue);
        Assert.Null(metadata.CustomData);
    }

    [Fact]
    public void CellMetadataBuilder_WithNullValues_AllowsNulls()
    {
        var metadata = CellMetadataBuilder.Create()
            .WithSource(null)
            .WithImportedAt(null)
            .WithOriginalValue(null)
            .WithCustomData(null)
            .Build();

        Assert.Null(metadata.Source);
        Assert.Null(metadata.ImportedAt);
        Assert.Null(metadata.OriginalValue);
        Assert.Null(metadata.CustomData);
    }
}
