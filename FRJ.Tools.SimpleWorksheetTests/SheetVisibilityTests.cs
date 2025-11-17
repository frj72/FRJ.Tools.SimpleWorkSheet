using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class SheetVisibilityTests
{
    [Fact]
    public void SetVisible_True_SetsVisible()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.SetVisible(true);
        
        Assert.True(sheet.IsVisible);
    }

    [Fact]
    public void SetVisible_False_SetsHidden()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.SetVisible(false);
        
        Assert.False(sheet.IsVisible);
    }

    [Fact]
    public void IsVisible_DefaultValue_IsTrue()
    {
        var sheet = new WorkSheet("TestSheet");
        
        Assert.True(sheet.IsVisible);
    }

    [Fact]
    public void SetVisible_CanToggle_UpdatesValue()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.SetVisible(false);
        Assert.False(sheet.IsVisible);
        
        sheet.SetVisible(true);
        Assert.True(sheet.IsVisible);
    }

    [Fact]
    public void Visibility_SaveAndLoad_PreservesVisibleState()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Test");
        sheet.SetVisible(true);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        Assert.Single(loadedWorkbook.Sheets);
        Assert.True(loadedWorkbook.Sheets.First().IsVisible);
    }

    [Fact]
    public void Visibility_SaveAndLoad_PreservesHiddenState()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Test");
        sheet.SetVisible(false);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        Assert.Single(loadedWorkbook.Sheets);
        Assert.False(loadedWorkbook.Sheets.First().IsVisible);
    }

    [Fact]
    public void Visibility_MultipleSheets_EachHasOwnState()
    {
        var sheet1 = new WorkSheet("Visible");
        sheet1.AddCell(new(0, 0), "Sheet 1");
        sheet1.SetVisible(true);
        
        var sheet2 = new WorkSheet("Hidden");
        sheet2.AddCell(new(0, 0), "Sheet 2");
        sheet2.SetVisible(false);
        
        var sheet3 = new WorkSheet("AlsoVisible");
        sheet3.AddCell(new(0, 0), "Sheet 3");
        
        var workbook = new WorkBook("Test", [sheet1, sheet2, sheet3]);
        var binary = SheetConverter.ToBinaryExcelFile(workbook);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        var sheets = loadedWorkbook.Sheets.ToList();
        Assert.Equal(3, sheets.Count);
        Assert.True(sheets[0].IsVisible);
        Assert.False(sheets[1].IsVisible);
        Assert.True(sheets[2].IsVisible);
    }

    [Fact]
    public void Visibility_WithoutSetting_RoundTripReturnsVisible()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Test");
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        Assert.True(loadedWorkbook.Sheets.First().IsVisible);
    }

    [Fact]
    public void Visibility_HiddenSheet_WithOtherFeatures_PreservesAll()
    {
        var sheet = new WorkSheet("DataSheet");
        sheet.AddCell(new(0, 0), "Calculation Data", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("FF0000")));
        sheet.SetVisible(false);
        sheet.SetTabColor("808080");
        sheet.FreezePanes(1, 0);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(binary);
        var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
        
        var loadedSheet = loadedWorkbook.Sheets.First();
        Assert.False(loadedSheet.IsVisible);
        Assert.Equal("808080", loadedSheet.TabColor);
        Assert.NotNull(loadedSheet.FrozenPane);
    }
}
