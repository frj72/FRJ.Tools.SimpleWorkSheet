namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellValueExtensions
{
    public static decimal AsDecimal(this CellValue value) => value.Value.AsDecimal();
    public static double AsDouble(this CellValue value) => value.Value.AsDouble();
    public static long AsLong(this CellValue value) => value.Value.AsLong();
    public static int AsInt(this CellValue value) => value.Value.AsInt();
    public static string AsString(this CellValue value) => value.Value.AsString();
    public static DateTime AsDateTime(this CellValue value) => value.Value.AsDateTime();
    public static DateTimeOffset AsDateTimeOffset(this CellValue value) => value.Value.AsDateTimeOffset();

    public static CellValueBasicType CellValueType(this CellValue cellValue)
    {
        if (cellValue.IsDecimal())
            return CellValueBasicType.FloatingPointNumber;
        if (cellValue.IsLong())
            return CellValueBasicType.IntegerNumber;
        if (cellValue.IsString())
            return CellValueBasicType.String;
        if (cellValue.IsDateTime() || cellValue.IsDateTimeOffset())
            return CellValueBasicType.DateType;
        return CellValueBasicType.String;
        
    }
    
}
