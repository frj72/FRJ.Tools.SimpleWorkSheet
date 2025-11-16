namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public class CellStyleBuilder
{
    private string? _fillColor;
    private CellFont? _font;
    private CellBorders? _borders;
    private string? _formatCode;
    private HorizontalAlignment? _horizontalAlignment;
    private VerticalAlignment? _verticalAlignment;
    private int? _textRotation;
    private bool? _wrapText;

    public CellStyleBuilder()
    {
        _fillColor = null;
        _font = null;
        _borders = null;
        _formatCode = null;
        _horizontalAlignment = null;
        _verticalAlignment = null;
        _textRotation = null;
        _wrapText = null;
    }

    public CellStyleBuilder(CellStyle? existingStyle)
    {
        if (existingStyle == null) return;
        _fillColor = existingStyle.FillColor;
        _font = existingStyle.Font;
        _borders = existingStyle.Borders;
        _formatCode = existingStyle.FormatCode;
        _horizontalAlignment = existingStyle.HorizontalAlignment;
        _verticalAlignment = existingStyle.VerticalAlignment;
        _textRotation = existingStyle.TextRotation;
        _wrapText = existingStyle.WrapText;
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

    public CellStyleBuilder WithHorizontalAlignment(HorizontalAlignment alignment)
    {
        _horizontalAlignment = alignment;
        return this;
    }

    public CellStyleBuilder WithVerticalAlignment(VerticalAlignment alignment)
    {
        _verticalAlignment = alignment;
        return this;
    }

    public CellStyleBuilder WithAlignment(HorizontalAlignment horizontal, VerticalAlignment vertical)
    {
        _horizontalAlignment = horizontal;
        _verticalAlignment = vertical;
        return this;
    }

    public CellStyleBuilder WithTextRotation(int degrees)
    {
        if (degrees is < -90 or > 90)
            throw new ArgumentException("Text rotation must be between -90 and 90 degrees", nameof(degrees));
        _textRotation = degrees;
        return this;
    }

    public CellStyleBuilder WithWrapText(bool wrap)
    {
        _wrapText = wrap;
        return this;
    }

    public CellStyle Build() => 
        CellStyle.Create(_fillColor, _font, _borders, _formatCode, _horizontalAlignment, _verticalAlignment, _textRotation, _wrapText);

    public static CellStyleBuilder Create() => new();

    public static CellStyleBuilder FromStyle(CellStyle style) => new(style);
}
