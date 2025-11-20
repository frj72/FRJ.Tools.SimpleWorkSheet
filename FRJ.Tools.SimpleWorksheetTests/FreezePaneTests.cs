using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class FreezePaneTests
{
    [Fact]
    public void FreezePanes_BothRowAndColumn_StoresCorrectly()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.FreezePanes(1, 2);
        
        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(1u, sheet.FrozenPane.Row);
        Assert.Equal(2u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezeRows_OnlyRow_StoresCorrectly()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.FreezeRows(3);
        
        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(3u, sheet.FrozenPane.Row);
        Assert.Equal(0u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezeColumns_OnlyColumn_StoresCorrectly()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.FreezeColumns(5);
        
        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(0u, sheet.FrozenPane.Row);
        Assert.Equal(5u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezePanes_NoFreeze_ReturnsNull()
    {
        var sheet = new WorkSheet("Test");
        
        Assert.Null(sheet.FrozenPane);
    }

    [Fact]
    public void FreezePanes_OverwriteExisting_UpdatesValue()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.FreezePanes(1, 1);
        sheet.FreezePanes(2, 3);
        
        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(2u, sheet.FrozenPane.Row);
        Assert.Equal(3u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezePanes_ZeroZero_SetsZeroZeroFreeze()
    {
        var sheet = new WorkSheet("Test");

        sheet.FreezePanes(0, 0);

        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(0u, sheet.FrozenPane.Row);
        Assert.Equal(0u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezePanes_FirstColumnOnly_FreezeColumn()
    {
        var sheet = new WorkSheet("Test");

        sheet.FreezePanes(0, 1);

        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(0u, sheet.FrozenPane.Row);
        Assert.Equal(1u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezePanes_FirstRowOnly_FreezeRow()
    {
        var sheet = new WorkSheet("Test");

        sheet.FreezePanes(1, 0);

        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(1u, sheet.FrozenPane.Row);
        Assert.Equal(0u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezePanes_BothRowsAndColumns_FreezesBoth()
    {
        var sheet = new WorkSheet("Test");

        sheet.FreezePanes(5, 5);

        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(5u, sheet.FrozenPane.Row);
        Assert.Equal(5u, sheet.FrozenPane.Column);
    }

    [Fact]
    public void FreezePanes_ExtremeValues_StoresValues()
    {
        var sheet = new WorkSheet("Test");

        sheet.FreezePanes(uint.MaxValue, uint.MaxValue);

        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(uint.MaxValue, sheet.FrozenPane.Row);
        Assert.Equal(uint.MaxValue, sheet.FrozenPane.Column);
    }

    private static readonly string[] SourceArrayHeaders = ["Column A", "Column B", "Column C", "Column D", "Column E"];

    [Fact]
    public void FreezePanes_CreatesValidExcelFile()
    {
        var sheet = new WorkSheet("FreezePanes");
        
        var headers = SourceArrayHeaders.Select(h => new CellValue(h));
        sheet.AddRow(0, 0, headers, configure: cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        
        for (uint row = 1; row < 20; row++)
        for (uint col = 0; col < 5; col++)
            sheet.AddCell(new(col, row), $"R{row}C{col}", null);

        sheet.FreezePanes(1, 0);
        
        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);
        
        Assert.NotEmpty(bytes);
        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void FreezePanes_MultipleScenarios_CreatesValidFiles()
    {
        var testCases = new[]
        {
            (row: 1u, col: 0u, name: "FreezeFirstRow"),
            (row: 0u, col: 1u, name: "FreezeFirstColumn"),
            (row: 1u, col: 1u, name: "FreezeBoth"),
            (row: 3u, col: 2u, name: "FreezeMultiple")
        };

        foreach (var (row, col, name) in testCases)
        {
            var sheet = new WorkSheet(name);
            
            for (uint r = 0; r < 10; r++)
            for (uint c = 0; c < 5; c++)
                sheet.AddCell(new(c, r), $"R{r}C{c}", null);

            sheet.FreezePanes(row, col);
            
            var workbook = new WorkBook("Test", [sheet]);
            var bytes = SheetConverter.ToBinaryExcelFile(workbook);
            
            Assert.NotEmpty(bytes);
        }
    }
}
