namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellColorExtensions
{
    public static bool IsValidColor(this string? hexValue) =>
        string.IsNullOrEmpty(hexValue) || hexValue is [ _, _, _, _, _, _] &&
        hexValue[..].All(c => char.IsDigit(c) || c is >= 'a' and <= 'f' || c is >= 'A' and <= 'F');
}
