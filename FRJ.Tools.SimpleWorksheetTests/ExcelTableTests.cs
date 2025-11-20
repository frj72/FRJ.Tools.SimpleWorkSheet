using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ExcelTableTests
{
    [Fact]
    public void AddTable_ValidParameters_AddsTable()
    {
        var sheet = new WorkSheet("TestSheet");
        var range = CellRange.FromBounds(0, 0, 2, 5);
        
        var table = sheet.AddTable("Table1", range);
        
        Assert.NotNull(table);
        Assert.Equal("Table1", table.Name);
        Assert.Equal(range, table.Range);
        Assert.True(table.ShowFilterButton);
        Assert.Single(sheet.Tables);
    }

    [Fact]
    public void AddTable_WithBounds_AddsTable()
    {
        var sheet = new WorkSheet("TestSheet");
        
        var table = sheet.AddTable("Table1", 0, 0, 2, 5);
        
        Assert.NotNull(table);
        Assert.Equal("Table1", table.Name);
        Assert.Single(sheet.Tables);
    }

    [Fact]
    public void AddTable_NoFilterButton_CreatesTableWithoutFilter()
    {
        var sheet = new WorkSheet("TestSheet");
        var range = CellRange.FromBounds(0, 0, 2, 5);
        
        var table = sheet.AddTable("Table1", range, showFilterButton: false);
        
        Assert.False(table.ShowFilterButton);
    }

    [Fact]
    public void AddTable_DuplicateName_ThrowsException()
    {
        var sheet = new WorkSheet("TestSheet");
        var range1 = CellRange.FromBounds(0, 0, 2, 5);
        var range2 = CellRange.FromBounds(5, 0, 7, 5);
        
        sheet.AddTable("Table1", range1);
        
        var ex = Assert.Throws<ArgumentException>(() => sheet.AddTable("Table1", range2));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void AddTable_OverlappingRange_ThrowsException()
    {
        var sheet = new WorkSheet("TestSheet");
        var range1 = CellRange.FromBounds(0, 0, 5, 5);
        var range2 = CellRange.FromBounds(3, 3, 8, 8);
        
        sheet.AddTable("Table1", range1);
        
        var ex = Assert.Throws<ArgumentException>(() => sheet.AddTable("Table2", range2));
        Assert.Contains("overlaps", ex.Message);
    }

    [Fact]
    public void ExcelTable_Constructor_EmptyName_ThrowsException()
    {
        var range = CellRange.FromBounds(0, 0, 2, 5);
        
        const string emptyName = "";
        var ex = Assert.Throws<ArgumentException>(() => new ExcelTable(emptyName, range, true, null));
        Assert.Contains("cannot be empty", ex.Message);
    }

    [Fact]
    public void ExcelTable_Constructor_SingleCell_ThrowsException()
    {
        var range = CellRange.FromBounds(0, 0, 0, 0);
        
        var ex = Assert.Throws<ArgumentException>(() => new ExcelTable("Table1", range, true, null));
        Assert.Contains("at least two cells", ex.Message);
    }

    [Fact]
    public void AddTable_SaveAndLoad_PreservesTable()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Header1", null);
        sheet.AddCell(new(1, 0), "Header2", null);
        sheet.AddCell(new(0, 1), "Data1", null);
        sheet.AddCell(new(1, 1), "Data2", null);
        
        var range = CellRange.FromBounds(0, 0, 1, 1);
        sheet.AddTable("TestTable", range);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        var loadedSheet = loadedWorkbook.Sheets.First();
        Assert.Single(loadedSheet.Tables);
        Assert.Equal("TestTable", loadedSheet.Tables[0].Name);
        Assert.Equal(range, loadedSheet.Tables[0].Range);
    }

    [Fact]
    public void AddTable_WithoutFilter_SaveAndLoad_PreservesFilterSetting()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Header", null);
        sheet.AddCell(new(0, 1), "Data", null);
        
        var range = CellRange.FromBounds(0, 0, 0, 1);
        sheet.AddTable("NoFilterTable", range, showFilterButton: false);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        var loadedSheet = loadedWorkbook.Sheets.First();
        Assert.False(loadedSheet.Tables[0].ShowFilterButton);
    }

    [Fact]
    public void AddTable_MultipleTables_AllPreserved()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Table1 Header", null);
        sheet.AddCell(new(0, 1), "Table1 Data", null);
        sheet.AddCell(new(5, 0), "Table2 Header", null);
        sheet.AddCell(new(5, 1), "Table2 Data", null);
        
        sheet.AddTable("Table1", 0, 0, 2, 3);
        sheet.AddTable("Table2", 5, 0, 7, 3);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        var loadedSheet = loadedWorkbook.Sheets.First();
        Assert.Equal(2, loadedSheet.Tables.Count);
        Assert.Contains(loadedSheet.Tables, t => t.Name == "Table1");
        Assert.Contains(loadedSheet.Tables, t => t.Name == "Table2");
    }

    [Fact]
    public void Tables_DefaultValue_IsEmpty()
    {
        var sheet = new WorkSheet("TestSheet");
        
        Assert.Empty(sheet.Tables);
    }
}
