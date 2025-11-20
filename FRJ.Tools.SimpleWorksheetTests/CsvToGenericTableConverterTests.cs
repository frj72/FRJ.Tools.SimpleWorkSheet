using FRJ.Tools.SimpleWorkSheet.Components.Import;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CsvToGenericTableConverterTests
{
    [Fact]
    public void Convert_EmptyCsv_ReturnsEmptyTable()
    {
        const string csv = "";
        
        var table = CsvToGenericTableConverter.Convert(csv);
        
        Assert.NotNull(table);
        Assert.Equal(0, table.ColumnCount);
        Assert.Equal(0, table.RowCount);
    }

    [Fact]
    public void Convert_WithHeader_CreatesTableWithHeaders()
    {
        const string csv = "name,age\nJohn,30";
        
        var table = CsvToGenericTableConverter.Convert(csv, hasHeader: true);
        
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(1, table.RowCount);
        Assert.Equal("name", table.Headers[0]);
        Assert.Equal("age", table.Headers[1]);
    }

    [Fact]
    public void Convert_WithoutHeader_GeneratesColumnNames()
    {
        const string csv = "John,30";
        
        var table = CsvToGenericTableConverter.Convert(csv, hasHeader: false);
        
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(1, table.RowCount);
        Assert.Equal("Column1", table.Headers[0]);
        Assert.Equal("Column2", table.Headers[1]);
    }

    [Fact]
    public void Convert_ParsesNumbers_AsDecimal()
    {
        const string csv = "value\n123\n456.78";
        
        var table = CsvToGenericTableConverter.Convert(csv);
        
        Assert.Equal(123m, table.GetValue(0, 0)?.Value.AsT0);
        Assert.Equal(456.78m, table.GetValue(0, 1)?.Value.AsT0);
    }

    [Fact]
    public void Convert_ParsesText_AsString()
    {
        const string csv = "name\nAlice\nBob";
        
        var table = CsvToGenericTableConverter.Convert(csv);
        
        Assert.Equal("Alice", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Bob", table.GetValue(0, 1)?.Value.AsT2);
    }

    [Fact]
    public void Convert_EmptyCells_StoresAsNull()
    {
        const string csv = "a,b,c\n1,,3";
        
        var table = CsvToGenericTableConverter.Convert(csv);
        
        Assert.NotNull(table.GetValue(0, 0));
        Assert.Null(table.GetValue(1, 0));
        Assert.NotNull(table.GetValue(2, 0));
    }

    [Fact]
    public void Convert_MultipleRows_CreatesAllRows()
    {
        const string csv = "name,score\nAlice,95\nBob,87\nCarol,92";
        
        var table = CsvToGenericTableConverter.Convert(csv);
        
        Assert.Equal(3, table.RowCount);
        Assert.Equal("Alice", table.GetValue("name", 0)?.Value.AsT2);
        Assert.Equal("Bob", table.GetValue("name", 1)?.Value.AsT2);
        Assert.Equal("Carol", table.GetValue("name", 2)?.Value.AsT2);
    }

    [Fact]
    public void ParseCellValue_Integer_ReturnsDecimal()
    {
        var value = CsvToGenericTableConverter.ParseCellValue("100");
        
        Assert.Equal(100m, value.Value.AsT0);
    }

    [Fact]
    public void ParseCellValue_Decimal_ReturnsDecimal()
    {
        var value = CsvToGenericTableConverter.ParseCellValue("99.99");
        
        Assert.Equal(99.99m, value.Value.AsT0);
    }

    [Fact]
    public void ParseCellValue_Date_ReturnsDateTime()
    {
        var value = CsvToGenericTableConverter.ParseCellValue("2024-01-15");
        
        Assert.True(value.Value.IsT3);
        var date = value.Value.AsT3;
        Assert.Equal(2024, date.Year);
        Assert.Equal(1, date.Month);
        Assert.Equal(15, date.Day);
    }

    [Fact]
    public void ParseCellValue_Text_ReturnsString()
    {
        var value = CsvToGenericTableConverter.ParseCellValue("Hello World");
        
        Assert.Equal("Hello World", value.Value.AsT2);
    }

    [Fact]
    public void ParseCellValue_NegativeNumber_ReturnsDecimal()
    {
        var value = CsvToGenericTableConverter.ParseCellValue("-42.5");
        
        Assert.Equal(-42.5m, value.Value.AsT0);
    }

    [Fact]
    public void Convert_QuotedFields_HandlesCorrectly()
    {
        const string csv = "name,description\nTest,\"Value with, comma\"";
        
        var table = CsvToGenericTableConverter.Convert(csv);
        
        Assert.Equal("Value with, comma", table.GetValue("description", 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_MixedTypes_ParsesEachCorrectly()
    {
        const string csv = "name,age,salary,active\nJohn,30,50000.50,true";
        
        var table = CsvToGenericTableConverter.Convert(csv);
        
        Assert.Equal("John", table.GetValue("name", 0)?.Value.AsT2);
        Assert.Equal(30m, table.GetValue("age", 0)?.Value.AsT0);
        Assert.Equal(50000.50m, table.GetValue("salary", 0)?.Value.AsT0);
        Assert.Equal("true", table.GetValue("active", 0)?.Value.AsT2);
    }
}
