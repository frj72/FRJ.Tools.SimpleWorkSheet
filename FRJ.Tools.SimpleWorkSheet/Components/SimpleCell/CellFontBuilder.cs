namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public class CellFontBuilder
{
    private int? _size;
    private string? _name;
    private string? _color;
    private bool _bold;
    private bool _italic;
    private bool _underline;
    private bool _strike;

    public CellFontBuilder()
    {
        _size = null;
        _name = null;
        _color = null;
        _bold = false;
        _italic = false;
        _underline = false;
        _strike = false;
    }

    public CellFontBuilder(CellFont? existingFont)
    {
        if (existingFont == null) return;
        _size = existingFont.Size;
        _name = existingFont.Name;
        _color = existingFont.Color;
        _bold = existingFont.Bold;
        _italic = existingFont.Italic;
        _underline = existingFont.Underline;
        _strike = existingFont.Strike;
    }

    public CellFontBuilder WithSize(int size)
    {
        _size = size;
        return this;
    }

    public CellFontBuilder WithName(string name)
    {
        _name = name;
        return this;
    }

    public CellFontBuilder WithColor(string color)
    {
        if (!color.IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(color));
        _color = color;
        return this;
    }

    public CellFontBuilder Bold(bool value = true)
    {
        _bold = value;
        return this;
    }

    public CellFontBuilder Italic(bool value = true)
    {
        _italic = value;
        return this;
    }

    public CellFontBuilder Underline(bool value = true)
    {
        _underline = value;
        return this;
    }

    public CellFontBuilder Strike(bool value = true)
    {
        _strike = value;
        return this;
    }

    public CellFont Build() => 
        CellFont.Create(_size, _name, _color, _bold, _italic, _underline, _strike);

    public static CellFontBuilder Create() => new();

    public static CellFontBuilder FromFont(CellFont font) => new(font);
}
