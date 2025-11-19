using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class GenericTableTests
{
    [Fact]
    public void Create_EmptyTable_ReturnsEmptyTable()
    {
        var table = GenericTable.Create();
        
        Assert.NotNull(table);
        Assert.Empty(table.Headers);
        Assert.Empty(table.Rows);
        Assert.Equal(0, table.ColumnCount);
        Assert.Equal(0, table.RowCount);
    }

    [Fact]
    public void Create_WithHeaders_ReturnsTableWithHeaders()
    {
        var table = GenericTable.Create("Name", "Age", "City");
        
        Assert.NotNull(table);
        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(0, table.RowCount);
        Assert.Equal("Name", table.Headers[0]);
        Assert.Equal("Age", table.Headers[1]);
        Assert.Equal("City", table.Headers[2]);
    }

    [Fact]
    public void AddHeader_ValidHeader_AddsHeader()
    {
        var table = GenericTable.Create();
        
        table.AddHeader("Column1");
        
        Assert.Single(table.Headers);
        Assert.Equal("Column1", table.Headers[0]);
    }

    [Fact]
    public void AddHeader_EmptyHeader_ThrowsArgumentException()
    {
        var table = GenericTable.Create();
        
        Assert.Throws<ArgumentException>(() => table.AddHeader(""));
    }

    [Fact]
    public void AddHeaders_ValidHeaders_AddsAllHeaders()
    {
        var table = GenericTable.Create();
        
        table.AddHeaders("A", "B", "C");
        
        Assert.Equal(3, table.ColumnCount);
        Assert.Equal("A", table.Headers[0]);
        Assert.Equal("B", table.Headers[1]);
        Assert.Equal("C", table.Headers[2]);
    }

    [Fact]
    public void AddRow_ValidRow_AddsRow()
    {
        var table = GenericTable.Create("Name", "Age");
        
        table.AddRow(new CellValue("John"), new CellValue(30));
        
        Assert.Single(table.Rows);
        Assert.Equal("John", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal(30, table.GetValue(1, 0)?.Value.AsT1);
    }

    [Fact]
    public void AddRow_WithList_AddsRow()
    {
        var table = GenericTable.Create("Name", "Age");
        
        var values = new List<CellValue?> { new("Jane"), new(25) };
        table.AddRow(values);
        
        Assert.Single(table.Rows);
        Assert.Equal("Jane", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal(25, table.GetValue(1, 0)?.Value.AsT1);
    }

    [Fact]
    public void AddRow_WrongColumnCount_ThrowsArgumentException()
    {
        var table = GenericTable.Create("Name", "Age");
        
        Assert.Throws<ArgumentException>(() => table.AddRow(new CellValue("John")));
    }

    [Fact]
    public void AddRow_WithNullValues_AcceptsNulls()
    {
        var table = GenericTable.Create("Name", "Age", "City");
        
        table.AddRow(new CellValue("John"), null, new CellValue("NYC"));
        
        Assert.Single(table.Rows);
        Assert.Equal("John", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Null(table.GetValue(1, 0));
        Assert.Equal("NYC", table.GetValue(2, 0)?.Value.AsT2);
    }

    [Fact]
    public void GetValue_ValidIndices_ReturnsValue()
    {
        var table = GenericTable.Create("A", "B");
        table.AddRow(new CellValue("X"), new CellValue("Y"));
        table.AddRow(new CellValue("P"), new CellValue("Q"));
        
        Assert.Equal("X", table.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Y", table.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("P", table.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal("Q", table.GetValue(1, 1)?.Value.AsT2);
    }

    [Fact]
    public void GetValue_ByColumnName_ReturnsValue()
    {
        var table = GenericTable.Create("Name", "Age");
        table.AddRow(new CellValue("Alice"), new CellValue(28));
        
        Assert.Equal("Alice", table.GetValue("Name", 0)?.Value.AsT2);
        Assert.Equal(28, table.GetValue("Age", 0)?.Value.AsT1);
    }

    [Fact]
    public void GetValue_InvalidColumnName_ThrowsArgumentException()
    {
        var table = GenericTable.Create("Name", "Age");
        table.AddRow(new CellValue("Bob"), new CellValue(35));
        
        Assert.Throws<ArgumentException>(() => table.GetValue("InvalidColumn", 0));
    }

    [Fact]
    public void GetValue_InvalidIndex_ThrowsArgumentOutOfRangeException()
    {
        var table = GenericTable.Create("A", "B");
        table.AddRow(new CellValue("X"), new CellValue("Y"));
        
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetValue(5, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetValue(0, 5));
    }

    [Fact]
    public void GetHeader_ValidIndex_ReturnsHeader()
    {
        var table = GenericTable.Create("First", "Second", "Third");
        
        Assert.Equal("First", table.GetHeader(0));
        Assert.Equal("Second", table.GetHeader(1));
        Assert.Equal("Third", table.GetHeader(2));
    }

    [Fact]
    public void GetHeader_InvalidIndex_ThrowsArgumentOutOfRangeException()
    {
        var table = GenericTable.Create("A");
        
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetHeader(5));
    }

    [Fact]
    public void GetColumnIndex_ValidColumn_ReturnsIndex()
    {
        var table = GenericTable.Create("Alpha", "Beta", "Gamma");
        
        Assert.Equal(0, table.GetColumnIndex("Alpha"));
        Assert.Equal(1, table.GetColumnIndex("Beta"));
        Assert.Equal(2, table.GetColumnIndex("Gamma"));
    }

    [Fact]
    public void GetColumnIndex_InvalidColumn_ThrowsArgumentException()
    {
        var table = GenericTable.Create("A", "B");
        
        Assert.Throws<ArgumentException>(() => table.GetColumnIndex("C"));
    }

    [Fact]
    public void HasColumn_ExistingColumn_ReturnsTrue()
    {
        var table = GenericTable.Create("Name", "Age");
        
        Assert.True(table.HasColumn("Name"));
        Assert.True(table.HasColumn("Age"));
    }

    [Fact]
    public void HasColumn_NonExistingColumn_ReturnsFalse()
    {
        var table = GenericTable.Create("Name", "Age");
        
        Assert.False(table.HasColumn("City"));
    }

    [Fact]
    public void Rows_ReturnsReadOnlyCollection()
    {
        var table = GenericTable.Create("A", "B");
        table.AddRow(new CellValue("X"), new CellValue("Y"));
        
        var rows = table.Rows;
        
        Assert.Single(rows);
        Assert.IsType<IReadOnlyList<IReadOnlyList<CellValue?>>>(rows, exactMatch: false);
    }

    [Fact]
    public void Headers_ReturnsReadOnlyCollection()
    {
        var table = GenericTable.Create("A", "B", "C");
        
        var headers = table.Headers;
        
        Assert.Equal(3, headers.Count);
        Assert.IsType<IReadOnlyList<string>>(headers, exactMatch: false);
    }

    [Fact]
    public void AddRow_EmptyTableWithNoHeaders_AddsRowWithoutValidation()
    {
        var table = GenericTable.Create();
        
        table.AddRow(new CellValue("A"), new CellValue("B"));
        
        Assert.Single(table.Rows);
    }
}
