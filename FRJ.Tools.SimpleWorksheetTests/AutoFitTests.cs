using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class AutoFitTests
{
    [Fact]
    public void AutoFitColumn_WithCells_SetsColumnWidth()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "Short", null);
        sheet.AddCell(new(0, 1), "Much longer text", null);
        sheet.AddCell(new(0, 2), "Mid", null);
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.True(sheet.ExplicitColumnWidths[0].IsT0);
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitColumn_EmptyColumn_SetsDefaultWidth()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(1, 0), "Some text", null);
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitColumn_WithNumbers_SetsColumnWidth()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), 123456789, null);
        sheet.AddCell(new(0, 1), 42, null);
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitColumn_WithDecimals_SetsColumnWidth()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), 123.456m, null);
        sheet.AddCell(new(0, 1), 3.14159m, null);
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitColumn_WithDates_SetsColumnWidth()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), DateTime.Now, null);
        sheet.AddCell(new(0, 1), new DateTime(2024, 12, 25), null);
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitColumn_WithBoldFont_ConsidersFontWeight()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "Regular text", null);
        sheet.AddCell(new(0, 1), "Bold text", cell => cell
            .WithFont(font => font.Bold()));
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitColumn_WithLargeFontSize_SetsWiderColumn()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "Normal", cell => cell
            .WithFont(font => font.WithSize(11)));
        sheet.AddCell(new(0, 1), "Large", cell => cell
            .WithFont(font => font.WithSize(24)));
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitAllColumns_MultipleColumns_SetsAllWidths()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "Column A short", null);
        sheet.AddCell(new(1, 0), "Column B much longer text", null);
        sheet.AddCell(new(2, 0), "C", null);
        
        sheet.AddCell(new(0, 1), "A2", null);
        sheet.AddCell(new(1, 1), "B2", null);
        sheet.AddCell(new(2, 1), "C2 longer", null);
        
        sheet.AutoFitAllColumns();
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(1));
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(2));
        
        var width0 = sheet.ExplicitColumnWidths[0].AsT0;
        var width1 = sheet.ExplicitColumnWidths[1].AsT0;
        var width2 = sheet.ExplicitColumnWidths[2].AsT0;
        
        Assert.True(width0 > 0);
        Assert.True(width1 > 0);
        Assert.True(width2 > 0);
        Assert.True(width1 > width0);
    }

    [Fact]
    public void AutoFitAllColumns_EmptySheet_NoError()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AutoFitAllColumns();
        
        Assert.Empty(sheet.ExplicitColumnWidths);
    }

    [Fact]
    public void AutoFitColumn_SparseRows_CalculatesCorrectly()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "Row 0", null);
        sheet.AddCell(new(0, 10), "Row 10 longer text", null);
        sheet.AddCell(new(0, 100), "Row 100", null);
        
        sheet.AutoFitColumn(0);
        
        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        var width = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.True(width > 0);
    }

    [Fact]
    public void AutoFitColumn_OverwritesPreviousWidth_SetsNewValue()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "Text", null);
        sheet.SetColumnWidth(0, 50.0);
        
        var oldWidth = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.Equal(50.0, oldWidth);
        
        sheet.AutoFitColumn(0);
        
        var newWidth = sheet.ExplicitColumnWidths[0].AsT0;
        Assert.NotEqual(50.0, newWidth);
    }

    [Fact]
    public void AutoFitAllColumns_MixedContent_HandlesAllTypes()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "String", null);
        sheet.AddCell(new(1, 0), 12345, null);
        sheet.AddCell(new(2, 0), 123.45m, null);
        sheet.AddCell(new(3, 0), DateTime.Now, null);
        
        sheet.AutoFitAllColumns();
        
        Assert.Equal(4, sheet.ExplicitColumnWidths.Count);
        foreach (var width in sheet.ExplicitColumnWidths.Values)
        {
            Assert.True(width.IsT0);
            Assert.True(width.AsT0 > 0);
        }
    }
}
