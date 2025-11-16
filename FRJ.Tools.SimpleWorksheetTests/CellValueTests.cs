using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellValueTests
{
    [Fact]
    public void CellValue_WithDecimalMaxValue_StoresCorrectly()
    {
        var value = new CellValue(decimal.MaxValue);

        Assert.True(value.IsDecimal());
        Assert.Equal(decimal.MaxValue, value.Value.AsT0);
    }

    [Fact]
    public void CellValue_WithDecimalMinValue_StoresCorrectly()
    {
        var value = new CellValue(decimal.MinValue);

        Assert.True(value.IsDecimal());
        Assert.Equal(decimal.MinValue, value.Value.AsT0);
    }

    [Fact]
    public void CellValue_WithZeroDecimal_StoresCorrectly()
    {
        var value = new CellValue(0m);

        Assert.True(value.IsDecimal());
        Assert.Equal(0m, value.Value.AsT0);
    }

    [Fact]
    public void CellValue_WithDateTimeMinValue_StoresCorrectly()
    {
        var minDate = DateTime.MinValue;
        var value = new CellValue(minDate);

        Assert.True(value.IsDateTime());
        Assert.Equal(minDate, value.Value.AsT3);
    }

    [Fact]
    public void CellValue_WithDateTimeMaxValue_StoresCorrectly()
    {
        var maxDate = DateTime.MaxValue;
        var value = new CellValue(maxDate);

        Assert.True(value.IsDateTime());
        Assert.Equal(maxDate, value.Value.AsT3);
    }

    [Fact]
    public void CellValue_WithLeapYearDate_StoresCorrectly()
    {
        var leapDate = new DateTime(2024, 2, 29);
        var value = new CellValue(leapDate);

        Assert.True(value.IsDateTime());
        Assert.Equal(leapDate, value.Value.AsT3);
    }

    [Fact]
    public void CellValue_WithEmptyString_StoresCorrectly()
    {
        var value = new CellValue("");

        Assert.True(value.IsString());
        Assert.Equal("", value.Value.AsT2);
    }

    [Fact]
    public void CellValue_WithWhitespaceOnlyString_StoresCorrectly()
    {
        var value = new CellValue("   ");

        Assert.True(value.IsString());
        Assert.Equal("   ", value.Value.AsT2);
    }

    [Fact]
    public void CellValue_WithVeryLongString_StoresCorrectly()
    {
        var longString = new string('A', 10000);
        var value = new CellValue(longString);

        Assert.True(value.IsString());
        Assert.Equal(longString, value.Value.AsT2);
        Assert.Equal(10000, value.Value.AsT2.Length);
    }

    [Fact]
    public void CellValue_ImplicitConversionFromInt_WorksCorrectly()
    {
        CellValue value = 42;

        Assert.True(value.IsLong());
        Assert.Equal(42L, value.Value.AsT1);
    }

    [Fact]
    public void CellValue_ImplicitConversionFromLong_WorksCorrectly()
    {
        CellValue value = 9999999999L;

        Assert.True(value.IsLong());
        Assert.Equal(9999999999L, value.Value.AsT1);
    }
}
