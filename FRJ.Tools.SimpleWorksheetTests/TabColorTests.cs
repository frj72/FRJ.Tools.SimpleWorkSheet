using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class TabColorTests
{
    [Fact]
    public void SetTabColor_ValidColor_SetsTabColor()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.SetTabColor("FF0000");
        
        Assert.Equal("FF0000", sheet.TabColor);
    }

    [Fact]
    public void SetTabColor_InvalidColor_ThrowsException()
    {
        var sheet = new WorkSheet("TestSheet");
        
        const string invalidColor = "ZZZZZZ";
        var ex = Assert.Throws<ArgumentException>(() => sheet.SetTabColor(invalidColor));
        
        Assert.Contains("Invalid color format", ex.Message);
    }

    [Fact]
    public void TabColor_DefaultValue_IsNull()
    {
        var sheet = new WorkSheet("TestSheet");
        
        Assert.Null(sheet.TabColor);
    }

    [Fact]
    public void SetTabColor_CanChangeColor_UpdatesValue()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.SetTabColor("FF0000");
        Assert.Equal("FF0000", sheet.TabColor);
        
        sheet.SetTabColor("00FF00");
        Assert.Equal("00FF00", sheet.TabColor);
    }

    [Fact]
    public void TabColor_SaveAndLoad_PreservesColor()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Test", null);
        sheet.SetTabColor("4472C4");
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        Assert.Single(loadedWorkbook.Sheets);
        Assert.Equal("4472C4", loadedWorkbook.Sheets.First().TabColor);
    }

    [Fact]
    public void TabColor_MultipleSheets_EachHasOwnColor()
    {
        var sheet1 = new WorkSheet("Sheet1");
        sheet1.AddCell(new(0, 0), "Sheet 1", null);
        sheet1.SetTabColor("FF0000");
        
        var sheet2 = new WorkSheet("Sheet2");
        sheet2.AddCell(new(0, 0), "Sheet 2", null);
        sheet2.SetTabColor("00FF00");
        
        var sheet3 = new WorkSheet("Sheet3");
        sheet3.AddCell(new(0, 0), "Sheet 3", null);
        
        var workbook = new WorkBook("Test", [sheet1, sheet2, sheet3]);
        var binary = SheetConverter.ToBinaryExcelFile(workbook);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        var sheets = loadedWorkbook.Sheets.ToList();
        Assert.Equal(3, sheets.Count);
        Assert.Equal("FF0000", sheets[0].TabColor);
        Assert.Equal("00FF00", sheets[1].TabColor);
        Assert.Null(sheets[2].TabColor);
    }

    [Theory]
    [InlineData("000000")]
    [InlineData("FFFFFF")]
    [InlineData("FF5733")]
    [InlineData("C70039")]
    [InlineData("900C3F")]
    [InlineData("581845")]
    public void SetTabColor_VariousValidColors_SetsCorrectly(string color)
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.SetTabColor(color);
        
        Assert.Equal(color, sheet.TabColor);
    }

    [Fact]
    public void TabColor_WithoutSetting_RoundTripReturnsNull()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Test", null);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        Assert.Null(loadedWorkbook.Sheets.First().TabColor);
    }
}
