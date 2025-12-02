using FRJ.Tools.SimpleWorkSheet.Components.Import;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CsvParserEdgeCaseTests
{
    [Fact]
    public void Convert_EscapedQuotes_HandlesCorrectly()
    {
        const string csv = "text\n\"Hello \"\"World\"\"\"";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("Hello \"World\"", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_NewlinesInQuotedFields_PreservesNewlines()
    {
        const string csv = "text\n\"Line 1\nLine 2\"";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("Line 1\nLine 2", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_EmptyQuotedFields_HandlesCorrectly()
    {
        const string csv = "a,b,c\n\"\",value,";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("value", table.GetValue(1, 0)?.Value.AsT2);
        Assert.Null(table.GetValue(2, 0));
    }

    [Fact]
    public void Convert_TrailingCommas_AddsEmptyFields()
    {
        const string csv = "a,b,c\n1,2,3,";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(4, table.ColumnCount);
        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("3", table.GetValue(2, 0)?.Value.AsT2);
        Assert.Null(table.GetValue(3, 0));
    }

    [Fact]
    public void Convert_LeadingCommas_AddsEmptyFields()
    {
        const string csv = "a,b,c\n,1,2,3";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(4, table.ColumnCount);
        Assert.Null(table.GetValue(0, 0));
        Assert.Equal("1", table.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(2, 0)?.Value.AsT2);
        Assert.Equal("3", table.GetValue(3, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_OnlyCommas_HandlesCorrectly()
    {
        const string csv = "a,b,c\n,,";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(3, table.ColumnCount);
        Assert.Null(table.GetValue(0, 0));
        Assert.Null(table.GetValue(1, 0));
        Assert.Null(table.GetValue(2, 0));
    }

    [Fact]
    public void Convert_SingleFieldNoHeader_HandlesCorrectly()
    {
        const string csv = "value";

        var table = CsvToGenericTableConverter.Convert(csv, hasHeader: false);

        Assert.Equal(1, table.ColumnCount);
        Assert.Equal(1, table.RowCount);
        Assert.Equal("Column1", table.Headers[0]);
        Assert.Equal("value", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_SingleFieldWithHeader_HandlesCorrectly()
    {
        const string csv = "name\nvalue";

        var table = CsvToGenericTableConverter.Convert(csv, hasHeader: true);

        Assert.Equal(1, table.ColumnCount);
        Assert.Equal(1, table.RowCount);
        Assert.Equal("name", table.Headers[0]);
        Assert.Equal("value", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_EmptyLines_AreIgnored()
    {
        const string csv = "a,b\n1,2\n\n3,4";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(2, table.RowCount);
        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("3", table.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal("4", table.GetValue(1, 1)?.Value.AsT2);
    }

    [Fact]
    public void Convert_CRLF_LineEndings_HandlesCorrectly()
    {
        const string csv = "a,b\r\n1,2\r\n3,4";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(2, table.RowCount);
        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(1, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_LF_LineEndings_HandlesCorrectly()
    {
        const string csv = "a,b\n1,2\n3,4";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(2, table.RowCount);
        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(1, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_CR_LineEndings_HandlesCorrectly()
    {
        const string csv = "a,b\r1,2\r3,4";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(2, table.RowCount);
        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(1, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_MixedLineEndings_HandlesCorrectly()
    {
        const string csv = "a,b\r\n1,2\n3,4\r5,6";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(3, table.RowCount);
        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("5", table.GetValue(0, 2)?.Value.AsT2);
    }

    [Fact]
    public void Convert_QuotedFieldsWithCommas_PreservesCommas()
    {
        const string csv = "name,city\n\"John Doe\",\"New York, NY\"";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("John Doe", table.GetValue("name", 0)?.Value.AsT2);
        Assert.Equal("New York, NY", table.GetValue("city", 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_QuotedFieldsWithQuotes_EscapesCorrectly()
    {
        const string csv = "text\n\"He said \"\"Hello\"\"\"";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("He said \"Hello\"", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_UnclosedQuotes_HandlesGracefully()
    {
        const string csv = "text\n\"Unclosed quote";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(1, table.RowCount);
        Assert.Equal("\"Unclosed quote", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_EmptyRows_AreIgnored()
    {
        const string csv = "a,b\n\n1,2\n\n";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(1, table.RowCount);
        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(1, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_HeaderOnly_NoDataRows()
    {
        const string csv = "name,age";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(0, table.RowCount);
        Assert.Equal("name", table.Headers[0]);
        Assert.Equal("age", table.Headers[1]);
    }

    [Fact]
    public void Convert_InconsistentColumnCount_FillsWithNulls()
    {
        const string csv = "a,b,c\n1,2\n3,4,5,6";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(4, table.ColumnCount);
        Assert.Equal(2, table.RowCount);

        Assert.Equal("1", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("2", table.GetValue(1, 0)?.Value.AsT2);
        Assert.Null(table.GetValue(2, 0));
        Assert.Null(table.GetValue(3, 0));
        Assert.Equal("3", table.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal("4", table.GetValue(1, 1)?.Value.AsT2);
        Assert.Equal("5", table.GetValue(2, 1)?.Value.AsT2);
        Assert.Equal("6", table.GetValue(3, 1)?.Value.AsT2);
    }

    [Fact]
    public void Convert_ScientificNotation_ParsesAsDecimal()
    {
        const string csv = "value\n1.23E-4\n5.67E+2";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(0.000123m, table.GetValue(0, 0)?.Value.AsT0);
        Assert.Equal(567m, table.GetValue(0, 1)?.Value.AsT0);
    }

    [Fact]
    public void Convert_DateTimeWithTime_ParsesCorrectly()
    {
        const string csv = "datetime\n2024-01-15 14:30:45";

        var table = CsvToGenericTableConverter.Convert(csv);

        var dateTime = table.GetValue(0, 0)?.Value.AsT3;
        Assert.Equal(2024, dateTime?.Year);
        Assert.Equal(1, dateTime?.Month);
        Assert.Equal(15, dateTime?.Day);
        Assert.Equal(14, dateTime?.Hour);
        Assert.Equal(30, dateTime?.Minute);
        Assert.Equal(45, dateTime?.Second);
    }

    [Fact]
    public void Convert_DateTimeOffset_ParsesCorrectly()
    {
        const string csv = "datetime\n2024-01-15T14:30:45+02:00";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.True(table.GetValue(0, 0)?.Value.IsT4);
    }

    [Fact]
    public void Convert_LargeNumbers_ParsesAsDecimal()
    {
        const string csv = "value\n12345678901234567890";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(12345678901234567890m, table.GetValue(0, 0)?.Value.AsT0);
    }

    [Fact]
    public void Convert_ZeroValues_ParsesCorrectly()
    {
        const string csv = "value\n0\n0.0";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(0m, table.GetValue(0, 0)?.Value.AsT0);
        Assert.Equal(0m, table.GetValue(0, 1)?.Value.AsT0);
    }

    [Fact]
    public void Convert_NegativeZero_ParsesCorrectly()
    {
        const string csv = "value\n-0";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(0m, table.GetValue(0, 0)?.Value.AsT0);
    }

    [Fact]
    public void Convert_UnicodeCharacters_HandlesCorrectly()
    {
        const string csv = "text\nHello 世界 🌍";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("Hello 世界 🌍", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_SpecialCharacters_HandlesCorrectly()
    {
        const string csv = "text\n!@#$%^&*()_+-=[]{}|;:,.<>?";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("!@#$%^&*()_+-=[]{}|;:,.<>?", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_TabsInUnquotedFields_HandlesCorrectly()
    {
        const string csv = "text\nHello\tWorld";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("Hello\tWorld", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_TabsInQuotedFields_PreservesTabs()
    {
        const string csv = "text\n\"Hello\tWorld\"";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("Hello\tWorld", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_BackslashEscapes_HandlesCorrectly()
    {
        const string csv = """
text
"Path: C:\\Program Files\\App"
""";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(@"Path: C:\\Program Files\\App", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_SingleQuoteInField_HandlesCorrectly()
    {
        const string csv = "text\nIt's working";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("It's working", table.GetValue(0, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_EmptyHeaderNames_HandlesCorrectly()
    {
        const string csv = ",b,\n1,2,3";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal("", table.Headers[0]);
        Assert.Equal("b", table.Headers[1]);
        Assert.Equal("", table.Headers[2]);
    }

    [Fact]
    public void Convert_WhitespaceInHeaders_PreservesWhitespace()
    {
        const string csv = " name , age \nJohn,30";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(" name ", table.Headers[0]);
        Assert.Equal(" age ", table.Headers[1]);
    }

    [Fact]
    public void Convert_WhitespaceInData_PreservesWhitespace()
    {
        const string csv = "name,age\n John , 30 ";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal(" John ", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal(" 30 ", table.GetValue(1, 0)?.Value.AsT2);
    }

    [Fact]
    public void Convert_OnlyWhitespace_ParsesAsString()
    {
        const string csv = "text\n   \n\t\t";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("   ", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("\t\t", table.GetValue(0, 1)?.Value.AsT2);
    }

    [Fact]
    public void Convert_QuotedWhitespace_PreservesWhitespace()
    {
        const string csv = "text\n\"   \"\n\"\t\t\"";

        var table = CsvToGenericTableConverter.Convert(csv);

        Assert.Equal("   ", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("\t\t", table.GetValue(0, 1)?.Value.AsT2);
    }
}