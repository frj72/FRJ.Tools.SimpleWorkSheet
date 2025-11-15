using System.Globalization;

namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellValueUnionExtensions
{
    extension(CellValueUnion value)
    {
        public decimal AsDecimal() =>
            value.Match(
                d => d,
                l => l,
                s => decimal.TryParse(s, out var result) ? result : decimal.Zero,
                dt => dt.Ticks,
                dto => dto.Ticks,
                _ => throw new InvalidOperationException("Cannot convert a formula to a decimal")
            );

        public double AsDouble() =>
            value.Match(
                d => (double) d,
                l => l,
                s => double.TryParse(s, out var result) ? result : 0.0,
                dt => dt.Ticks,
                dto => dto.Ticks,
                _ => throw new InvalidOperationException("Cannot convert a formula to a double")
            );

        public long AsLong() =>
            value.Match(
                d => (long) d,
                l => l,
                s => long.TryParse(s, out var result) ? result : 0L,
                dt => dt.Ticks,
                dto => dto.Ticks,
                _ => throw new InvalidOperationException("Cannot convert a formula to a long")
            );

        public int AsInt() =>
            value.Match(
                d => (int) d,
                l => (int) l,
                s => int.TryParse(s, out var result) ? result : 0,
                dt => (int) dt.Ticks,
                dto => (int) dto.Ticks,
                _ => throw new InvalidOperationException("Cannot convert a formula to an int")
            );

        public string AsString() =>
            value.Match<string>(
                d => d.ToString(CultureInfo.InvariantCulture),
                l => l.ToString(),
                s => s,
                dt => dt.ToString("O"),
                dto => dto.ToString("O"),
                cf => cf.Value
            );

        public DateTime AsDateTime() =>
            value.Match(
                d => new((long) d),
                l => new(l),
                s => DateTime.TryParse(s, null, DateTimeStyles.RoundtripKind, out var result) ? result : DateTime.MinValue,
                dt => dt,
                dto => dto.DateTime,
                _ => throw new InvalidOperationException("Cannot convert a formula to a DateTime")
            );

        public DateTimeOffset AsDateTimeOffset() =>
            value.Match(
                d => new DateTime((long) d),
                l => new DateTime(l),
                s => DateTimeOffset.TryParse(s, null, DateTimeStyles.RoundtripKind, out var result) ? result : DateTimeOffset.MinValue,
                dt => new(dt),
                dto => dto,
                _ => throw new InvalidOperationException("Cannot convert a formula to a DateTimeOffset")
            );
    }
}