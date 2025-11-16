namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellColorExtensions
{
    extension(string? hexValue)
    {
        public bool IsValidColor() =>
            string.IsNullOrEmpty(hexValue) || 
            hexValue is [ _, _, _, _, _, _] or [ _, _, _, _, _, _, _, _] &&
            hexValue[..].All(c => char.IsDigit(c) || c is >= 'a' and <= 'f' || c is >= 'A' and <= 'F');

        public string ToArgbColor()
        {
            if (string.IsNullOrEmpty(hexValue))
                return "FF000000";
            return hexValue.Length switch
            {
                8 => hexValue,
                6 => "FF" + hexValue,
                _ => throw new ArgumentException($"Invalid color format: {hexValue}. Expected 6 or 8 characters.",
                    nameof(hexValue))
            };
        }
    }
}
