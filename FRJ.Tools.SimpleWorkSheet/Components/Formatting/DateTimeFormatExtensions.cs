// ReSharper disable UnusedMember.Global
namespace FRJ.Tools.SimpleWorkSheet.Components.Formatting;

public static class DateTimeFormatExtensions
{
    extension(DateFormat dateFormat)
    {
        public string ToFormatString() => dateFormat switch
        {
            DateFormat.IsoDateTime => "yyyy-MM-dd HH:mm:ss",
            DateFormat.IsoDate => "yyyy-MM-dd",
            DateFormat.DateTime => "dd/MM/yyyy HH:mm:ss",
            DateFormat.DateOnly => "dd/MM/yyyy",
            DateFormat.TimeOnly => "HH:mm:ss",
            _ => "yyyy-MM-dd HH:mm:ss"
        };

        public string ToExcelFormatString() => dateFormat switch
        {
            DateFormat.IsoDateTime => "yyyy-mm-dd hh:mm:ss",
            DateFormat.IsoDate => "yyyy-mm-dd",
            DateFormat.DateTime => "dd/mm/yyyy hh:mm:ss",
            DateFormat.DateOnly => "dd/mm/yyyy",
            DateFormat.TimeOnly => "hh:mm:ss",
            _ => "yyyy-mm-dd hh:mm:ss"
        };

        public uint ToExcelFormatId() => dateFormat switch
        {
            DateFormat.IsoDateTime => 164,
            DateFormat.IsoDate => 165,
            DateFormat.DateTime => 166,
            DateFormat.DateOnly => 167,
            DateFormat.TimeOnly => 168,
            _ => throw new ArgumentOutOfRangeException(nameof(dateFormat), dateFormat, null)
        };
    }
}