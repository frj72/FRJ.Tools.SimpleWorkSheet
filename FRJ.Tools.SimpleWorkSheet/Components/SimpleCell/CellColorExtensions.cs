namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellColorExtensions
{
    public static bool IsValidColor(this string? hexValue) =>
        string.IsNullOrEmpty(hexValue) || 
        (hexValue is [ _, _, _, _, _, _] || hexValue is [ _, _, _, _, _, _, _, _]) &&
        hexValue[..].All(c => char.IsDigit(c) || c is >= 'a' and <= 'f' || c is >= 'A' and <= 'F');

    public static string ToArgbColor(this string? hexValue)
    {
        if (string.IsNullOrEmpty(hexValue))
            return "FF000000";
        if (hexValue.Length == 8)
            return hexValue;
        if (hexValue.Length == 6)
            return "FF" + hexValue;
        throw new ArgumentException($"Invalid color format: {hexValue}. Expected 6 or 8 characters.", nameof(hexValue));
    }
}
