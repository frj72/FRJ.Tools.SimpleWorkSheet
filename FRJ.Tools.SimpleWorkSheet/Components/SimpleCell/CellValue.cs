global using CellValueUnion = OneOf.OneOf<decimal, long, string, System.DateTime, System.DateTimeOffset, FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.CellFormula>;
namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public class CellValue
{
    public CellValueUnion Value { get; }
    public CellValue(CellValueUnion cellValue) => Value = cellValue; 
    public CellValue(decimal value) => Value = value;
    public CellValue(double value) => Value = (decimal) value;
    public CellValue(long value) => Value = value;
    public CellValue(int value) => Value = value;
    public CellValue(string value) => Value = value;
    public CellValue(DateTime value) => Value = value;
    public CellValue(DateTimeOffset value) => Value = value;
    public CellValue(CellFormula value) => Value = value;

    public bool IsDecimal() => Value.IsT0;
    public bool IsLong() => Value.IsT1;
    public bool IsString() => Value.IsT2;
    public bool IsDateTime() => Value.IsT3;
    public bool IsDateTimeOffset() => Value.IsT4;
    public bool IsCellFormula() => Value.IsT5;
    
    public static implicit operator CellValue(string val) => new(val);
    public static implicit operator CellValue(decimal val) => new(val);
    public static implicit operator CellValue(double val) => new(val);
    public static implicit operator CellValue(long val) => new(val);
    public static implicit operator CellValue(int val) => new(val);
    public static implicit operator CellValue(DateTime val) => new(val);
    public static implicit operator CellValue(DateTimeOffset val) => new(val);
    public static implicit operator CellValue(CellFormula val) => new(val);
}
