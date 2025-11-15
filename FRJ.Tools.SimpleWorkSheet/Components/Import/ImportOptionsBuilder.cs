using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class ImportOptionsBuilder
{
    private bool _preserveOriginalValue = true;
    private CellStyle? _defaultStyle;
    private Dictionary<int, Func<string, CellValue>>? _columnParsers;
    private string? _sourceIdentifier;
    private bool _trimWhitespace = true;
    private bool _skipEmptyRows;
    private Dictionary<string, object>? _customMetadata;

    public ImportOptionsBuilder()
    {
    }

    public ImportOptionsBuilder(ImportOptions? existingOptions)
    {
        if (existingOptions == null) return;
        _preserveOriginalValue = existingOptions.PreserveOriginalValue;
        _defaultStyle = existingOptions.DefaultStyle;
        _columnParsers = existingOptions.ColumnParsers;
        _sourceIdentifier = existingOptions.SourceIdentifier;
        _trimWhitespace = existingOptions.TrimWhitespace;
        _skipEmptyRows = existingOptions.SkipEmptyRows;
        _customMetadata = existingOptions.CustomMetadata != null
            ? new(existingOptions.CustomMetadata)
            : null;
    }

    public ImportOptionsBuilder WithPreserveOriginalValue(bool preserve)
    {
        _preserveOriginalValue = preserve;
        return this;
    }

    public ImportOptionsBuilder WithDefaultStyle(CellStyle? style)
    {
        _defaultStyle = style;
        return this;
    }

    public ImportOptionsBuilder WithDefaultStyle(Action<CellStyleBuilder> configure)
    {
        var builder = new CellStyleBuilder(_defaultStyle);
        configure(builder);
        _defaultStyle = builder.Build();
        return this;
    }

    public ImportOptionsBuilder WithColumnParser(int columnIndex, Func<string, CellValue> parser)
    {
        _columnParsers ??= new();
        _columnParsers[columnIndex] = parser;
        return this;
    }

    public ImportOptionsBuilder WithSourceIdentifier(string? sourceIdentifier)
    {
        _sourceIdentifier = sourceIdentifier;
        return this;
    }

    public ImportOptionsBuilder WithTrimWhitespace(bool trim)
    {
        _trimWhitespace = trim;
        return this;
    }

    public ImportOptionsBuilder WithSkipEmptyRows(bool skip)
    {
        _skipEmptyRows = skip;
        return this;
    }

    public ImportOptionsBuilder WithCustomMetadata(string key, object value)
    {
        _customMetadata ??= new();
        _customMetadata[key] = value;
        return this;
    }

    public ImportOptions Build() => ImportOptions.Create(
        _preserveOriginalValue,
        _defaultStyle,
        _columnParsers,
        _sourceIdentifier,
        _trimWhitespace,
        _skipEmptyRows,
        _customMetadata);

    public static ImportOptionsBuilder Create() => new();

    public static ImportOptionsBuilder FromOptions(ImportOptions options) => new(options);
}
