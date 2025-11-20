namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public record CellMetadata
{
    public string? Source { get; init; }
    public DateTime? ImportedAt { get; init; }
    public string? OriginalValue { get; init; }
    public Dictionary<string, object>? CustomData { get; init; }

    public static CellMetadata Create(string? source = null, DateTime? importedAt = null, string? originalValue = null,
        Dictionary<string, object>? customData = null)
        => new()
        {
            Source = source,
            ImportedAt = importedAt,
            OriginalValue = originalValue,
            CustomData = customData
        };
}