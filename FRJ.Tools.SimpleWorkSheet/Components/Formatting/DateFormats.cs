namespace FRJ.Tools.SimpleWorkSheet.Components.Formatting;

public enum DateFormat
{
    IsoDateTime,
    IsoDate,
    DateTime,
    DateOnly,
    TimeOnly,
}

public static class DateTimeFormatExtensions
{
    public static string ToFormatString(this DateFormat dateFormat)
    {
        return dateFormat switch
        {
            DateFormat.IsoDateTime => "yyyy-MM-dd HH:mm:ss",
            DateFormat.IsoDate => "yyyy-MM-dd",
            DateFormat.DateTime => "dd/MM/yyyy HH:mm:ss",
            DateFormat.DateOnly => "dd/MM/yyyy",
            DateFormat.TimeOnly => "HH:mm:ss",
            _ => "yyyy-MM-dd HH:mm:ss"
        };
    }

    public static string ToExcelFormatString(this DateFormat dateFormat)
    {
        return dateFormat switch
        {
            DateFormat.IsoDateTime => "yyyy-mm-dd hh:mm:ss",
            DateFormat.IsoDate => "yyyy-mm-dd",
            DateFormat.DateTime => "dd/mm/yyyy hh:mm:ss",
            DateFormat.DateOnly => "dd/mm/yyyy",
            DateFormat.TimeOnly => "hh:mm:ss",
            _ => "yyyy-mm-dd hh:mm:ss"
        };
    }

    public static uint ToExcelFormatId(this DateFormat dateFormat)
    {
        return dateFormat switch
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

public enum NumberFormat
{
    Integer,
    Float2,
    Float3,
    Float4,
}

public static class NumberFormatExtensions
{
    public static string ToFormatString(this NumberFormat numberFormat)
    {
        return numberFormat switch
        {
            NumberFormat.Integer => "0",
            NumberFormat.Float2 => "F2",
            NumberFormat.Float3 => "F3",
            NumberFormat.Float4 => "F4",
            _ => "F"
        };
    }

    public static string ToExcelFormatString(this NumberFormat numberFormat)
    {
        return numberFormat switch
        {
            NumberFormat.Integer => "0",
            NumberFormat.Float2 => "0.00",
            NumberFormat.Float3 => "0.000",
            NumberFormat.Float4 => "0.0000",
            _ => "0.00"
        };
    }
    
    public static uint ToExcelFormatId(this NumberFormat numberFormat)
    {
        return numberFormat switch
        {
            NumberFormat.Integer => 169,
            NumberFormat.Float2 => 170,
            NumberFormat.Float3 => 171,
            NumberFormat.Float4 => 172,
            _ => throw new ArgumentOutOfRangeException(nameof(numberFormat), numberFormat, null)
        };
    }
}
