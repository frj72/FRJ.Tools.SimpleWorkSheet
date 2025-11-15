using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellMetadataTests
{
    [Fact]
    public void CellMetadata_Create_WithAllParameters_SetsAllProperties()
    {
        const string source = "csv";
        var importedAt = DateTime.UtcNow;
        const string originalValue = "raw_value";
        var customData = new Dictionary<string, object> { { "key1", "value1" } };

        var metadata = CellMetadata.Create(source, importedAt, originalValue, customData);

        Assert.Equal(source, metadata.Source);
        Assert.Equal(importedAt, metadata.ImportedAt);
        Assert.Equal(originalValue, metadata.OriginalValue);
        Assert.Equal(customData, metadata.CustomData);
    }

    [Fact]
    public void CellMetadata_Create_WithNullParameters_SetsNullProperties()
    {
        var metadata = CellMetadata.Create();

        Assert.Null(metadata.Source);
        Assert.Null(metadata.ImportedAt);
        Assert.Null(metadata.OriginalValue);
        Assert.Null(metadata.CustomData);
    }

    [Fact]
    public void CellMetadata_Create_WithNoParameters_SetsNullProperties()
    {
        var metadata = CellMetadata.Create();

        Assert.Null(metadata.Source);
        Assert.Null(metadata.ImportedAt);
        Assert.Null(metadata.OriginalValue);
        Assert.Null(metadata.CustomData);
    }

    [Fact]
    public void CellMetadata_RecordEquality_WithSameValues_AreEqual()
    {
        const string source = "json";
        var importedAt = new DateTime(2025, 1, 1);
        var metadata1 = CellMetadata.Create(source, importedAt, "original");
        var metadata2 = CellMetadata.Create(source, importedAt, "original");

        Assert.Equal(metadata1, metadata2);
    }

    [Fact]
    public void CellMetadata_RecordEquality_WithDifferentValues_AreNotEqual()
    {
        var metadata1 = CellMetadata.Create("csv");
        var metadata2 = CellMetadata.Create("json");

        Assert.NotEqual(metadata1, metadata2);
    }

    [Fact]
    public void CellMetadata_WithExpression_UpdatesProperty()
    {
        var originalMetadata = CellMetadata.Create("csv");
        var importedAt = DateTime.UtcNow;
        
        var updatedMetadata = originalMetadata with { ImportedAt = importedAt };

        Assert.Equal("csv", updatedMetadata.Source);
        Assert.Equal(importedAt, updatedMetadata.ImportedAt);
    }

    [Fact]
    public void CellMetadata_CustomData_CanStoreMultipleValues()
    {
        var customData = new Dictionary<string, object>
        {
            { "column_index", 5 },
            { "row_index", 10 },
            { "validator", "email" }
        };

        var metadata = CellMetadata.Create("csv", null, null, customData);

        Assert.NotNull(metadata.CustomData);
        Assert.Equal(5, metadata.CustomData["column_index"]);
        Assert.Equal(10, metadata.CustomData["row_index"]);
        Assert.Equal("email", metadata.CustomData["validator"]);
    }
}
