namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public class CellStyleBuilder
{
    private string? _fillColor;
    private CellFont? _font;
    private CellBorders? _borders;
    private string? _formatCode;

    public CellStyleBuilder()
    {
        _fillColor = null;
        _font = null;
        _borders = null;
        _formatCode = null;
    }

    public CellStyleBuilder(CellStyle? existingStyle)
    {
        if (existingStyle == null) return;
        _fillColor = existingStyle.FillColor;
        _font = existingStyle.Font;
        _borders = existingStyle.Borders;
        _formatCode = existingStyle.FormatCode;
    }

    public CellStyleBuilder WithFillColor(string? fillColor)
    {
        if (fillColor != null && !fillColor.IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(fillColor));
        _fillColor = fillColor;
        return this;
    }

    public CellStyleBuilder WithFont(CellFont? font)
    {
        if (!font.HasValidColor())
            throw new ArgumentException("Invalid font color format", nameof(font));
        _font = font;
        return this;
    }

    public CellStyleBuilder WithFont(Action<CellFontBuilder> configure)
    {
        var builder = new CellFontBuilder(_font);
        configure(builder);
        _font = builder.Build();
        return this;
    }

    public CellStyleBuilder WithBorders(CellBorders? borders)
    {
        if (!borders.HasValidColors())
            throw new ArgumentException("Invalid border color format", nameof(borders));
        _borders = borders;
        return this;
    }

    public CellStyleBuilder WithFormatCode(string? formatCode)
    {
        _formatCode = formatCode;
        return this;
    }

    public CellStyle Build() => 
        CellStyle.Create(_fillColor, _font, _borders, _formatCode);

    public static CellStyleBuilder Create() => new();

    public static CellStyleBuilder FromStyle(CellStyle style) => new(style);
}
