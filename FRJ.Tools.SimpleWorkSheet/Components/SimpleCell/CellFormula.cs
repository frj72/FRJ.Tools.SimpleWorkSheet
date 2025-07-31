namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public struct CellFormula(string value)
{
    public string Value { get; } = value;
    public static implicit operator CellFormula(string val) => new(val);
}