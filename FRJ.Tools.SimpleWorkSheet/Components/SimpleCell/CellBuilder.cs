using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public class CellBuilder
{
    private CellValue _value;
    private CellStyle _style;
    private CellMetadata? _metadata;
    private CellHyperlink? _hyperlink;

    public CellBuilder()
    {
        _value = string.Empty;
        _style = WorkSheetDefaults.DefaultCellStyle;
        _metadata = null;
        _hyperlink = null;
    }

    public CellBuilder(Cell existingCell)
    {
        _value = existingCell.Value;
        _style = existingCell.Style;
        _metadata = existingCell.Metadata;
        _hyperlink = existingCell.Hyperlink;
    }

    public CellBuilder WithValue(CellValue value)
    {
        _value = value;
        return this;
    }

    public CellBuilder WithColor(string color)
    {
        if (!color.IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(color));
        _style = _style.WithFillColor(color);
        return this;
    }

    public CellBuilder WithFont(CellFont font)
    {
        if (!font.HasValidColor())
            throw new ArgumentException("Invalid font color format", nameof(font));
        _style = _style.WithFont(font);
        return this;
    }

    public CellBuilder WithFont(Action<CellFontBuilder> configure)
    {
        var builder = new CellFontBuilder(_style.Font);
        configure(builder);
        _style = _style.WithFont(builder.Build());
        return this;
    }

    public CellBuilder WithBorders(CellBorders borders)
    {
        if (!borders.HasValidColors())
            throw new ArgumentException("Invalid border color format", nameof(borders));
        _style = _style.WithBorders(borders);
        return this;
    }

    public CellBuilder WithFormatCode(string formatCode)
    {
        _style = _style.WithFormatCode(formatCode);
        return this;
    }

    public CellBuilder WithStyle(Action<CellStyleBuilder> configure)
    {
        var builder = new CellStyleBuilder(_style);
        configure(builder);
        _style = builder.Build();
        return this;
    }

    public CellBuilder WithStyle(CellStyle style)
    {
        if (!style.HasValidColors())
            throw new ArgumentException("Invalid style colors", nameof(style));
        _style = style;
        return this;
    }

    public CellBuilder WithMetadata(Action<CellMetadataBuilder> configure)
    {
        var builder = new CellMetadataBuilder(_metadata);
        configure(builder);
        _metadata = builder.Build();
        return this;
    }

    public CellBuilder WithMetadata(CellMetadata? metadata)
    {
        _metadata = metadata;
        return this;
    }

    public CellBuilder FromSource(string source)
    {
        _metadata ??= CellMetadata.Create();
        var builder = new CellMetadataBuilder(_metadata);
        builder.WithSource(source);
        _metadata = builder.Build();
        return this;
    }

    public CellBuilder WithHyperlink(string url, string? tooltip)
    {
        _hyperlink = new(url, tooltip);
        return this;
    }

    public Cell Build() => new(_value, _style, _metadata) { Hyperlink = _hyperlink };

    public static CellBuilder Create() => new();

    public static CellBuilder FromValue(CellValue value)
    {
        var builder = new CellBuilder
        {
            _value = value
        };
        return builder;
    }

    public static CellBuilder FromCell(Cell cell) => new(cell);
}
