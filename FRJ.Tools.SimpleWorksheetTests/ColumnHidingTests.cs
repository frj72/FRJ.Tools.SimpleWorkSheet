using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ColumnHidingTests
{
    [Fact]
    public void HideColumn_SingleColumn_MarksAsHidden()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(1, 0), "B1", null);
        
        sheet.HideColumn(1);
        
        Assert.True(sheet.HiddenColumns.ContainsKey(1));
        Assert.True(sheet.HiddenColumns[1]);
    }

    [Fact]
    public void ShowColumn_PreviouslyHidden_MarksAsVisible()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideColumn(1);
        
        sheet.ShowColumn(1);
        
        Assert.False(sheet.HiddenColumns.ContainsKey(1));
    }

    [Fact]
    public void SetColumnHidden_True_HidesColumn()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetColumnHidden(2, true);
        
        Assert.True(sheet.HiddenColumns.ContainsKey(2));
        Assert.True(sheet.HiddenColumns[2]);
    }

    [Fact]
    public void SetColumnHidden_False_ShowsColumn()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideColumn(2);
        
        sheet.SetColumnHidden(2, false);
        
        Assert.False(sheet.HiddenColumns.ContainsKey(2));
    }

    [Fact]
    public void HideColumns_MultipleColumns_HidesAll()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.HideColumns(1, 3, 5);
        
        Assert.True(sheet.HiddenColumns.ContainsKey(1));
        Assert.True(sheet.HiddenColumns.ContainsKey(3));
        Assert.True(sheet.HiddenColumns.ContainsKey(5));
    }

    [Fact]
    public void ShowColumns_MultipleColumns_ShowsAll()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideColumns(1, 2, 3);
        
        sheet.ShowColumns(1, 2, 3);
        
        Assert.False(sheet.HiddenColumns.ContainsKey(1));
        Assert.False(sheet.HiddenColumns.ContainsKey(2));
        Assert.False(sheet.HiddenColumns.ContainsKey(3));
    }

    [Fact]
    public void SetColumnWidth_CellWidthHidden_HidesColumn()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetColumnWidth(1, CellWidth.Hidden);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(1));
        Assert.True(sheet.ExplicitColumnWidths[1].IsT1);
        Assert.Equal(CellWidth.Hidden, sheet.ExplicitColumnWidths[1].AsT1);
    }

    [Fact]
    public void HiddenColumn_RoundTrip_PreservesHiddenState()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(1, 0), "B1", null);
        sheet.AddCell(new(2, 0), "C1", null);
        
        sheet.HideColumn(1);
        
        var tempFile = Path.GetTempFileName();
        var workbook = new WorkBook("Test", [sheet]);
        workbook.SaveToFile(tempFile);
        
        var readWorkbook = WorkBookReader.LoadFromFile(tempFile);
        var readSheet = readWorkbook.Sheets.First();
        
        Assert.True(readSheet.HiddenColumns.ContainsKey(1));
        Assert.True(readSheet.HiddenColumns[1]);
        Assert.False(readSheet.HiddenColumns.ContainsKey(0));
        Assert.False(readSheet.HiddenColumns.ContainsKey(2));
        
        File.Delete(tempFile);
    }

    [Fact]
    public void HiddenColumnViaEnum_RoundTrip_PreservesHiddenState()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(1, 0), "B1", null);
        
        sheet.SetColumnWidth(1, CellWidth.Hidden);
        
        var tempFile = Path.GetTempFileName();
        var workbook = new WorkBook("Test", [sheet]);
        workbook.SaveToFile(tempFile);
        
        var readWorkbook = WorkBookReader.LoadFromFile(tempFile);
        var readSheet = readWorkbook.Sheets.First();
        
        Assert.True(readSheet.HiddenColumns.ContainsKey(1));
        Assert.True(readSheet.HiddenColumns[1]);
        
        File.Delete(tempFile);
    }

    [Fact]
    public void HiddenColumn_WithExplicitWidth_BothPreserved()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(1, 0), "B1", null);
        
        sheet.SetColumnWidth(1, 25.0);
        sheet.HideColumn(1);
        
        var tempFile = Path.GetTempFileName();
        var workbook = new WorkBook("Test", [sheet]);
        workbook.SaveToFile(tempFile);
        
        var readWorkbook = WorkBookReader.LoadFromFile(tempFile);
        var readSheet = readWorkbook.Sheets.First();
        
        Assert.True(readSheet.HiddenColumns.ContainsKey(1));
        Assert.True(readSheet.ExplicitColumnWidths.ContainsKey(1));
        Assert.Equal(25.0, readSheet.ExplicitColumnWidths[1].AsT0);
        
        File.Delete(tempFile);
    }

    [Fact]
    public void HideColumn_AlreadyHidden_NoError()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideColumn(1);
        
        sheet.HideColumn(1);
        
        Assert.True(sheet.HiddenColumns.ContainsKey(1));
    }

    [Fact]
    public void ShowColumn_NotHidden_NoError()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.ShowColumn(1);
        
        Assert.False(sheet.HiddenColumns.ContainsKey(1));
    }
}
