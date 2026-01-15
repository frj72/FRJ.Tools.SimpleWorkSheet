// ReSharper disable UnusedMember.Global
namespace FRJ.Tools.SimpleWorkSheet.Components.Formatting;

public static class NumberFormatExtensions
{
    extension(NumberFormat numberFormat)
    {
        public string ToFormatString() => numberFormat switch
        {
            NumberFormat.Integer => "0",
            NumberFormat.Float2 => "F2",
            NumberFormat.Float3 => "F3",
            NumberFormat.Float4 => "F4",
            _ => "F"
        };

        public string ToExcelFormatString() => numberFormat switch
        {
            NumberFormat.Integer => "0",
            NumberFormat.Float2 => "0.00",
            NumberFormat.Float3 => "0.000",
            NumberFormat.Float4 => "0.0000",
            _ => "0.00"
        };

        public uint ToExcelFormatId() => numberFormat switch
        {
            NumberFormat.Integer => 169,
            NumberFormat.Float2 => 170,
            NumberFormat.Float3 => 171,
            NumberFormat.Float4 => 172,
            _ => throw new ArgumentOutOfRangeException(nameof(numberFormat), numberFormat, null)
        };
    }
}
