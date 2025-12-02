using System.Globalization;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellValueUnionExtensionsTests
{
    [Fact]
    public void AsDecimal_WithDecimalValue_ReturnsDecimal()
    {
        var cellValue = new CellValue(123.45m);
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(123.45m, result);
    }
    
    [Fact]
    public void AsDecimal_WithLongValue_ReturnsDecimalFromLong()
    {
        var cellValue = new CellValue(100L);
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(100m, result);
    }
    
    [Fact]
    public void AsDecimal_WithValidStringValue_ParsesDecimal()
    {
        var cellValue = new CellValue("456.78");
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(456.78m, result);
    }
    
    [Fact]
    public void AsDecimal_WithInvalidStringValue_ReturnsZero()
    {
        var cellValue = new CellValue("not a number");
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(decimal.Zero, result);
    }
    
    [Fact]
    public void AsDecimal_WithEmptyString_ReturnsZero()
    {
        var cellValue = new CellValue("");
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(decimal.Zero, result);
    }
    
    [Fact]
    public void AsDecimal_WithDateTimeValue_ReturnsTicks()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0);
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(dateTime.Ticks, result);
    }
    
    [Fact]
    public void AsDecimal_WithDateTimeOffsetValue_ReturnsTicks()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(dateTimeOffset.Ticks, result);
    }
    
    [Fact]
    public void AsDecimal_WithFormula_ThrowsInvalidOperationException()
    {
        var cellValue = new CellValue(new CellFormula("=SUM(A1:A10)"));
        
        var ex = Assert.Throws<InvalidOperationException>(() => cellValue.Value.AsDecimal());
        
        Assert.Contains("Cannot convert a formula to a decimal", ex.Message);
    }
    
    [Fact]
    public void AsDouble_WithDecimalValue_ReturnsDouble()
    {
        var cellValue = new CellValue(123.45m);
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(123.45, result, precision: 10);
    }
    
    [Fact]
    public void AsDouble_WithLongValue_ReturnsDoubleFromLong()
    {
        var cellValue = new CellValue(100L);
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(100.0, result);
    }
    
    [Fact]
    public void AsDouble_WithValidStringValue_ParsesDouble()
    {
        var cellValue = new CellValue("456.78");
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(456.78, result, precision: 10);
    }
    
    [Fact]
    public void AsDouble_WithInvalidStringValue_ReturnsZero()
    {
        var cellValue = new CellValue("not a number");
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(0.0, result);
    }
    
    [Fact]
    public void AsDouble_WithEmptyString_ReturnsZero()
    {
        var cellValue = new CellValue("");
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(0.0, result);
    }
    
    [Fact]
    public void AsDouble_WithDateTimeValue_ReturnsTicks()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0);
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(dateTime.Ticks, result);
    }
    
    [Fact]
    public void AsDouble_WithDateTimeOffsetValue_ReturnsTicks()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(dateTimeOffset.Ticks, result);
    }
    
    [Fact]
    public void AsDouble_WithFormula_ThrowsInvalidOperationException()
    {
        var cellValue = new CellValue(new CellFormula("=SUM(A1:A10)"));
        
        var ex = Assert.Throws<InvalidOperationException>(() => cellValue.Value.AsDouble());
        
        Assert.Contains("Cannot convert a formula to a double", ex.Message);
    }
    
    [Fact]
    public void AsLong_WithDecimalValue_ReturnsLong()
    {
        var cellValue = new CellValue(123.45m);
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(123L, result);
    }
    
    [Fact]
    public void AsLong_WithLongValue_ReturnsLong()
    {
        var cellValue = new CellValue(100L);
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(100L, result);
    }
    
    [Fact]
    public void AsLong_WithValidStringValue_ParsesLong()
    {
        var cellValue = new CellValue("456");
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(456L, result);
    }
    
    [Fact]
    public void AsLong_WithInvalidStringValue_ReturnsZero()
    {
        var cellValue = new CellValue("not a number");
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(0L, result);
    }
    
    [Fact]
    public void AsLong_WithEmptyString_ReturnsZero()
    {
        var cellValue = new CellValue("");
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(0L, result);
    }
    
    [Fact]
    public void AsLong_WithDateTimeValue_ReturnsTicks()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0);
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(dateTime.Ticks, result);
    }
    
    [Fact]
    public void AsLong_WithDateTimeOffsetValue_ReturnsTicks()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(dateTimeOffset.Ticks, result);
    }
    
    [Fact]
    public void AsLong_WithFormula_ThrowsInvalidOperationException()
    {
        var cellValue = new CellValue(new CellFormula("=SUM(A1:A10)"));
        
        var ex = Assert.Throws<InvalidOperationException>(() => cellValue.Value.AsLong());
        
        Assert.Contains("Cannot convert a formula to a long", ex.Message);
    }
    
    [Fact]
    public void AsInt_WithDecimalValue_ReturnsInt()
    {
        var cellValue = new CellValue(123.45m);
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(123, result);
    }
    
    [Fact]
    public void AsInt_WithLongValue_ReturnsInt()
    {
        var cellValue = new CellValue(100L);
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(100, result);
    }
    
    [Fact]
    public void AsInt_WithValidStringValue_ParsesInt()
    {
        var cellValue = new CellValue("456");
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(456, result);
    }
    
    [Fact]
    public void AsInt_WithInvalidStringValue_ReturnsZero()
    {
        var cellValue = new CellValue("not a number");
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(0, result);
    }
    
    [Fact]
    public void AsInt_WithEmptyString_ReturnsZero()
    {
        var cellValue = new CellValue("");
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(0, result);
    }
    
    [Fact]
    public void AsInt_WithDateTimeValue_ReturnsTicks()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0);
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal((int)dateTime.Ticks, result);
    }
    
    [Fact]
    public void AsInt_WithDateTimeOffsetValue_ReturnsTicks()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal((int)dateTimeOffset.Ticks, result);
    }
    
    [Fact]
    public void AsInt_WithFormula_ThrowsInvalidOperationException()
    {
        var cellValue = new CellValue(new CellFormula("=SUM(A1:A10)"));
        
        var ex = Assert.Throws<InvalidOperationException>(() => cellValue.Value.AsInt());
        
        Assert.Contains("Cannot convert a formula to an int", ex.Message);
    }
    
    [Fact]
    public void AsString_WithDecimalValue_ReturnsInvariantCultureString()
    {
        var cellValue = new CellValue(123.45m);
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal("123.45", result);
    }
    
    [Fact]
    public void AsString_WithLongValue_ReturnsString()
    {
        var cellValue = new CellValue(100L);
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal("100", result);
    }
    
    [Fact]
    public void AsString_WithStringValue_ReturnsString()
    {
        var cellValue = new CellValue("test string");
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal("test string", result);
    }
    
    [Fact]
    public void AsString_WithEmptyString_ReturnsEmptyString()
    {
        var cellValue = new CellValue("");
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal("", result);
    }
    
    [Fact]
    public void AsString_WithDateTimeValue_ReturnsRoundtripFormat()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0, DateTimeKind.Utc);
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal(dateTime.ToString("O"), result);
    }
    
    [Fact]
    public void AsString_WithDateTimeOffsetValue_ReturnsRoundtripFormat()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal(dateTimeOffset.ToString("O"), result);
    }
    
    [Fact]
    public void AsString_WithFormula_ReturnsFormulaValue()
    {
        var cellValue = new CellValue(new CellFormula("=SUM(A1:A10)"));
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal("=SUM(A1:A10)", result);
    }
    
    [Fact]
    public void AsDateTime_WithDecimalValue_CreatesDateTimeFromTicks()
    {
        const decimal ticks = 638400000000000000m;
        var cellValue = new CellValue(ticks);
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(new DateTime((long)ticks), result);
    }
    
    [Fact]
    public void AsDateTime_WithLongValue_CreatesDateTimeFromTicks()
    {
        const long ticks = 638400000000000000L;
        var cellValue = new CellValue(ticks);
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(new DateTime(ticks), result);
    }
    
    [Fact]
    public void AsDateTime_WithValidStringValue_ParsesDateTime()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0, DateTimeKind.Utc);
        var cellValue = new CellValue(dateTime.ToString("O"));
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(dateTime, result);
    }
    
    [Fact]
    public void AsDateTime_WithInvalidStringValue_ReturnsMinValue()
    {
        var cellValue = new CellValue("not a date");
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(DateTime.MinValue, result);
    }
    
    [Fact]
    public void AsDateTime_WithEmptyString_ReturnsMinValue()
    {
        var cellValue = new CellValue("");
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(DateTime.MinValue, result);
    }
    
    [Fact]
    public void AsDateTime_WithDateTimeValue_ReturnsDateTime()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0);
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(dateTime, result);
    }
    
    [Fact]
    public void AsDateTime_WithDateTimeOffsetValue_ReturnsDateTime()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(dateTimeOffset.DateTime, result);
    }
    
    [Fact]
    public void AsDateTime_WithFormula_ThrowsInvalidOperationException()
    {
        var cellValue = new CellValue(new CellFormula("=TODAY()"));
        
        var ex = Assert.Throws<InvalidOperationException>(() => cellValue.Value.AsDateTime());
        
        Assert.Contains("Cannot convert a formula to a DateTime", ex.Message);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithDecimalValue_CreatesDateTimeOffsetFromTicks()
    {
        const decimal ticks = 638400000000000000m;
        var cellValue = new CellValue(ticks);
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(new DateTimeOffset(new DateTime((long)ticks)), result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithLongValue_CreatesDateTimeOffsetFromTicks()
    {
        const long ticks = 638400000000000000L;
        var cellValue = new CellValue(ticks);
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(new DateTimeOffset(new DateTime(ticks)), result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithValidStringValue_ParsesDateTimeOffset()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset.ToString("O"));
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(dateTimeOffset, result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithInvalidStringValue_ReturnsMinValue()
    {
        var cellValue = new CellValue("not a date");
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(DateTimeOffset.MinValue, result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithEmptyString_ReturnsMinValue()
    {
        var cellValue = new CellValue("");
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(DateTimeOffset.MinValue, result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithDateTimeValue_ReturnsDateTimeOffset()
    {
        var dateTime = new DateTime(2024, 1, 1, 12, 0, 0);
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(new DateTimeOffset(dateTime), result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithDateTimeOffsetValue_ReturnsDateTimeOffset()
    {
        var dateTimeOffset = new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero);
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(dateTimeOffset, result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithFormula_ThrowsInvalidOperationException()
    {
        var cellValue = new CellValue(new CellFormula("=NOW()"));
        
        var ex = Assert.Throws<InvalidOperationException>(() => cellValue.Value.AsDateTimeOffset());
        
        Assert.Contains("Cannot convert a formula to a DateTimeOffset", ex.Message);
    }
    
    [Fact]
    public void AsDecimal_WithNegativeDecimal_PreservesSign()
    {
        var cellValue = new CellValue(-123.45m);
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(-123.45m, result);
    }
    
    [Fact]
    public void AsDecimal_WithNegativeString_ParsesNegative()
    {
        var cellValue = new CellValue("-456.78");
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(-456.78m, result);
    }
    
    [Fact]
    public void AsDouble_WithScientificNotationString_ParsesCorrectly()
    {
        var cellValue = new CellValue("1.23E+10");
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(1.23E+10, result);
    }
    
    [Fact]
    public void AsLong_WithNegativeLong_PreservesSign()
    {
        var cellValue = new CellValue(-100L);
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(-100L, result);
    }
    
    [Fact]
    public void AsInt_WithNegativeInt_PreservesSign()
    {
        var cellValue = new CellValue(-100);
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(-100, result);
    }
    
    [Fact]
    public void AsString_WithMaxDecimal_FormatsCorrectly()
    {
        var cellValue = new CellValue(decimal.MaxValue);
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal(decimal.MaxValue.ToString(CultureInfo.InvariantCulture), result);
    }
    
    [Fact]
    public void AsDecimal_WithMaxLong_ConvertsCorrectly()
    {
        var cellValue = new CellValue(long.MaxValue);
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(long.MaxValue, result);
    }
    
    [Fact]
    public void AsDateTime_WithMinDateTime_HandlesCorrectly()
    {
        var dateTime = DateTime.MinValue;
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(DateTime.MinValue, result);
    }
    
    [Fact]
    public void AsDateTime_WithMaxDateTime_HandlesCorrectly()
    {
        var dateTime = DateTime.MaxValue;
        var cellValue = new CellValue(dateTime);
        
        var result = cellValue.Value.AsDateTime();
        
        Assert.Equal(DateTime.MaxValue, result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithMinDateTimeOffset_HandlesCorrectly()
    {
        var dateTimeOffset = DateTimeOffset.MinValue;
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(DateTimeOffset.MinValue, result);
    }
    
    [Fact]
    public void AsDateTimeOffset_WithMaxDateTimeOffset_HandlesCorrectly()
    {
        var dateTimeOffset = DateTimeOffset.MaxValue;
        var cellValue = new CellValue(dateTimeOffset);
        
        var result = cellValue.Value.AsDateTimeOffset();
        
        Assert.Equal(DateTimeOffset.MaxValue, result);
    }
    
    [Fact]
    public void AsString_WithUnicodeCharacters_PreservesUnicode()
    {
        var cellValue = new CellValue("Hello 世界 🌍");
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal("Hello 世界 🌍", result);
    }
    
    [Fact]
    public void AsDecimal_WithWhitespaceString_ReturnsZero()
    {
        var cellValue = new CellValue("   ");
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(decimal.Zero, result);
    }
    
    [Fact]
    public void AsDouble_WithWhitespaceString_ReturnsZero()
    {
        var cellValue = new CellValue("   ");
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(0.0, result);
    }
    
    [Fact]
    public void AsLong_WithWhitespaceString_ReturnsZero()
    {
        var cellValue = new CellValue("   ");
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(0L, result);
    }
    
    [Fact]
    public void AsInt_WithWhitespaceString_ReturnsZero()
    {
        var cellValue = new CellValue("   ");
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(0, result);
    }
    
    [Fact]
    public void AsString_WithWhitespaceString_PreservesWhitespace()
    {
        var cellValue = new CellValue("   ");
        
        var result = cellValue.Value.AsString();
        
        Assert.Equal("   ", result);
    }
    
    [Fact]
    public void AsDecimal_WithZeroDecimal_ReturnsZero()
    {
        var cellValue = new CellValue(0m);
        
        var result = cellValue.Value.AsDecimal();
        
        Assert.Equal(0m, result);
    }
    
    [Fact]
    public void AsDouble_WithZeroDouble_ReturnsZero()
    {
        var cellValue = new CellValue(0.0);
        
        var result = cellValue.Value.AsDouble();
        
        Assert.Equal(0.0, result);
    }
    
    [Fact]
    public void AsLong_WithZeroLong_ReturnsZero()
    {
        var cellValue = new CellValue(0L);
        
        var result = cellValue.Value.AsLong();
        
        Assert.Equal(0L, result);
    }
    
    [Fact]
    public void AsInt_WithZeroInt_ReturnsZero()
    {
        var cellValue = new CellValue(0);
        
        var result = cellValue.Value.AsInt();
        
        Assert.Equal(0, result);
    }
}
