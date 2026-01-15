// ReSharper disable UnusedMember.Global
namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellValueExtensions
{
    extension(CellValue value)
    {
        public decimal AsDecimal() => value.Value.AsDecimal();
        public double AsDouble() => value.Value.AsDouble();
        public long AsLong() => value.Value.AsLong();
        public int AsInt() => value.Value.AsInt();
        public string AsString() => value.Value.AsString();
        public DateTime AsDateTime() => value.Value.AsDateTime();
        public DateTimeOffset AsDateTimeOffset() => value.Value.AsDateTimeOffset();

        public CellValueBasicType CellValueType()
        {
            if (value.IsDecimal())
                return CellValueBasicType.FloatingPointNumber;
            if (value.IsLong())
                return CellValueBasicType.IntegerNumber;
            if (value.IsString())
                return CellValueBasicType.String;
            if (value.IsDateTime() || value.IsDateTimeOffset())
                return CellValueBasicType.DateType;
            return CellValueBasicType.String;
        }
    }
}
