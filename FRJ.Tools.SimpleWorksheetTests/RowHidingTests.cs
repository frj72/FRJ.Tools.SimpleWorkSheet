using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class RowHidingTests
{
    [Fact]
    public void HideRow_SingleRow_MarksAsHidden()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(0, 1), "A2", null);
        
        sheet.HideRow(1);
        
        Assert.True(sheet.HiddenRows.ContainsKey(1));
        Assert.True(sheet.HiddenRows[1]);
    }

    [Fact]
    public void ShowRow_PreviouslyHidden_MarksAsVisible()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideRow(1);
        
        sheet.ShowRow(1);
        
        Assert.False(sheet.HiddenRows.ContainsKey(1));
    }

    [Fact]
    public void SetRowHidden_True_HidesRow()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetRowHidden(2, true);
        
        Assert.True(sheet.HiddenRows.ContainsKey(2));
        Assert.True(sheet.HiddenRows[2]);
    }

    [Fact]
    public void SetRowHidden_False_ShowsRow()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideRow(2);
        
        sheet.SetRowHidden(2, false);
        
        Assert.False(sheet.HiddenRows.ContainsKey(2));
    }

    [Fact]
    public void HideRows_MultipleRows_HidesAll()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.HideRows(1, 3, 5);
        
        Assert.True(sheet.HiddenRows.ContainsKey(1));
        Assert.True(sheet.HiddenRows.ContainsKey(3));
        Assert.True(sheet.HiddenRows.ContainsKey(5));
    }

    [Fact]
    public void ShowRows_MultipleRows_ShowsAll()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideRows(1, 2, 3);
        
        sheet.ShowRows(1, 2, 3);
        
        Assert.False(sheet.HiddenRows.ContainsKey(1));
        Assert.False(sheet.HiddenRows.ContainsKey(2));
        Assert.False(sheet.HiddenRows.ContainsKey(3));
    }

    [Fact]
    public void SetRowHeight_RowHeightHidden_HidesRow()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetRowHeight(1, RowHeight.Hidden);
        
        Assert.True(sheet.ExplicitRowHeights.ContainsKey(1));
        Assert.True(sheet.ExplicitRowHeights[1].IsT1);
        Assert.Equal(RowHeight.Hidden, sheet.ExplicitRowHeights[1].AsT1);
    }

    [Fact]
    public void HiddenRow_RoundTrip_PreservesHiddenState()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(0, 1), "A2", null);
        sheet.AddCell(new(0, 2), "A3", null);
        
        sheet.HideRow(1);
        
        var tempFile = Path.GetTempFileName();
        var workbook = new WorkBook("Test", [sheet]);
        workbook.SaveToFile(tempFile);
        
        var readWorkbook = WorkBookReader.LoadFromFile(tempFile);
        var readSheet = readWorkbook.Sheets.First();
        
        Assert.True(readSheet.HiddenRows.ContainsKey(1));
        Assert.True(readSheet.HiddenRows[1]);
        Assert.False(readSheet.HiddenRows.ContainsKey(0));
        Assert.False(readSheet.HiddenRows.ContainsKey(2));
        
        File.Delete(tempFile);
    }

    [Fact]
    public void HiddenRowViaEnum_RoundTrip_PreservesHiddenState()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(0, 1), "A2", null);
        
        sheet.SetRowHeight(1, RowHeight.Hidden);
        
        var tempFile = Path.GetTempFileName();
        var workbook = new WorkBook("Test", [sheet]);
        workbook.SaveToFile(tempFile);
        
        var readWorkbook = WorkBookReader.LoadFromFile(tempFile);
        var readSheet = readWorkbook.Sheets.First();
        
        Assert.True(readSheet.HiddenRows.ContainsKey(1));
        Assert.True(readSheet.HiddenRows[1]);
        
        File.Delete(tempFile);
    }

    [Fact]
    public void HiddenRow_WithExplicitHeight_BothPreserved()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "A1", null);
        sheet.AddCell(new(0, 1), "A2", null);
        
        sheet.SetRowHeight(1, 30.0);
        sheet.HideRow(1);
        
        var tempFile = Path.GetTempFileName();
        var workbook = new WorkBook("Test", [sheet]);
        workbook.SaveToFile(tempFile);
        
        var readWorkbook = WorkBookReader.LoadFromFile(tempFile);
        var readSheet = readWorkbook.Sheets.First();
        
        Assert.True(readSheet.HiddenRows.ContainsKey(1));
        Assert.True(readSheet.ExplicitRowHeights.ContainsKey(1));
        Assert.Equal(30.0, readSheet.ExplicitRowHeights[1].AsT0);
        
        File.Delete(tempFile);
    }

    [Fact]
    public void HideRow_AlreadyHidden_NoError()
    {
        var sheet = new WorkSheet("Test");
        sheet.HideRow(1);
        
        sheet.HideRow(1);
        
        Assert.True(sheet.HiddenRows.ContainsKey(1));
    }

    [Fact]
    public void ShowRow_NotHidden_NoError()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.ShowRow(1);
        
        Assert.False(sheet.HiddenRows.ContainsKey(1));
    }
}
