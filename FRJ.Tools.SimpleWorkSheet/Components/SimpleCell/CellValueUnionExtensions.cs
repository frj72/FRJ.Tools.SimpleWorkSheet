using System.Globalization;

namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellValueUnionExtensions
{
    public static decimal AsDecimal(this CellValueUnion value)
    {
        return value.Match(
            d => d,
            l => l,
            s => decimal.TryParse(s, out var result) ? result : decimal.Zero,
            dt => dt.Ticks,
            dto => dto.Ticks,
            cf => throw new InvalidOperationException("Cannot convert a formula to a decimal")
        );
    }
    public static double AsDouble(this CellValueUnion value)
    {
        return value.Match(
            d => (double) d,
            l => l,
            s => double.TryParse(s, out var result) ? result : 0.0,
            dt => dt.Ticks,
            dto => dto.Ticks,
            cf => throw new InvalidOperationException("Cannot convert a formula to a double")
        );
    }
    
    public static long AsLong(this CellValueUnion value)
    {
        return value.Match(
            d => (long) d,
            l => l,
            s => long.TryParse(s, out var result) ? result : 0L,
            dt => dt.Ticks,
            dto => dto.Ticks,
            cf => throw new InvalidOperationException("Cannot convert a formula to a long")
        );
    }
    
    public static int AsInt(this CellValueUnion value)
    {
        return value.Match(
            d => (int) d,
            l => (int) l,
            s => int.TryParse(s, out var result) ? result : 0,
            dt => (int) dt.Ticks,
            dto => (int) dto.Ticks,
            cf => throw new InvalidOperationException("Cannot convert a formula to an int")
        );
    }
    
    public static string AsString(this CellValueUnion value)
    {
        return value.Match<string>(
            d => d.ToString(CultureInfo.InvariantCulture),
            l => l.ToString(),
            s => s,
            dt => dt.ToString("O"),
            dto => dto.ToString("O"),
            cf => cf.Value
        );
    }
    
    public static DateTime AsDateTime(this CellValueUnion value)
    {
        return value.Match(
            d => new((long) d),
            l => new(l),
            s => DateTime.TryParse(s, null, DateTimeStyles.RoundtripKind, out var result) ? result : DateTime.MinValue,
            dt => dt,
            dto => dto.DateTime,
            cf => throw new InvalidOperationException("Cannot convert a formula to a DateTime")
        );
    }
    
    public static DateTimeOffset AsDateTimeOffset(this CellValueUnion value)
    {
        return value.Match(
            d => new DateTime((long) d),
            l => new DateTime(l),
            s => DateTimeOffset.TryParse(s, null, DateTimeStyles.RoundtripKind, out var result) ? result : DateTimeOffset.MinValue,
            dt => new(dt),
            dto => dto,
            cf => throw new InvalidOperationException("Cannot convert a formula to a DateTimeOffset")
        );
    }
}