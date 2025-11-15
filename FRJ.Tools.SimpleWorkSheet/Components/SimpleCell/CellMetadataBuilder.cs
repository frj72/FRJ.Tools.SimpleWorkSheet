namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public class CellMetadataBuilder
{
    private string? _source;
    private DateTime? _importedAt;
    private string? _originalValue;
    private Dictionary<string, object>? _customData;

    public CellMetadataBuilder()
    {
        _source = null;
        _importedAt = null;
        _originalValue = null;
        _customData = null;
    }

    public CellMetadataBuilder(CellMetadata? existingMetadata)
    {
        if (existingMetadata == null) return;
        _source = existingMetadata.Source;
        _importedAt = existingMetadata.ImportedAt;
        _originalValue = existingMetadata.OriginalValue;
        _customData = existingMetadata.CustomData != null 
            ? new(existingMetadata.CustomData) 
            : null;
    }

    public CellMetadataBuilder WithSource(string? source)
    {
        _source = source;
        return this;
    }

    public CellMetadataBuilder WithImportedAt(DateTime? importedAt)
    {
        _importedAt = importedAt;
        return this;
    }

    public CellMetadataBuilder WithOriginalValue(string? originalValue)
    {
        _originalValue = originalValue;
        return this;
    }

    public CellMetadataBuilder WithCustomData(Dictionary<string, object>? customData)
    {
        _customData = customData;
        return this;
    }

    public CellMetadataBuilder AddCustomData(string key, object value)
    {
        _customData ??= new();
        _customData[key] = value;
        return this;
    }

    public CellMetadata Build() => 
        CellMetadata.Create(_source, _importedAt, _originalValue, _customData);

    public static CellMetadataBuilder Create() => new();

    public static CellMetadataBuilder FromMetadata(CellMetadata metadata) => new(metadata);
}
