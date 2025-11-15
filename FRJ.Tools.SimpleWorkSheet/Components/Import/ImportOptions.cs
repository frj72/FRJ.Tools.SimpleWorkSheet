using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public record ImportOptions
{
    public bool PreserveOriginalValue { get; init; } = true;
    public CellStyle? DefaultStyle { get; init; }
    public Dictionary<int, Func<string, CellValue>>? ColumnParsers { get; init; }
    public string? SourceIdentifier { get; init; }
    public bool TrimWhitespace { get; init; } = true;
    public bool SkipEmptyRows { get; init; }
    public Dictionary<string, object>? CustomMetadata { get; init; }

    public static ImportOptions Create(
        bool preserveOriginalValue = true,
        CellStyle? defaultStyle = null,
        Dictionary<int, Func<string, CellValue>>? columnParsers = null,
        string? sourceIdentifier = null,
        bool trimWhitespace = true,
        bool skipEmptyRows = false,
        Dictionary<string, object>? customMetadata = null) => new()
    {
        PreserveOriginalValue = preserveOriginalValue,
        DefaultStyle = defaultStyle,
        ColumnParsers = columnParsers,
        SourceIdentifier = sourceIdentifier,
        TrimWhitespace = trimWhitespace,
        SkipEmptyRows = skipEmptyRows,
        CustomMetadata = customMetadata
    };
}
