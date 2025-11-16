using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class WorkBookTests
{
    [Fact]
    public void WorkBook_WithSingleSheet_CreatesInstance()
    {
        var sheet = new WorkSheet("Sheet1");

        var workBook = new WorkBook("Test", [sheet]);

        Assert.Single(workBook.Sheets);
        Assert.Same(sheet, workBook.Sheets.First());
    }

    [Fact]
    public void WorkBook_WithTenSheets_PreservesOrder()
    {
        var sheets = Enumerable.Range(1, 10)
            .Select(i => new WorkSheet($"Sheet{i}"))
            .ToList();

        var workBook = new WorkBook("Test", sheets);

        Assert.Equal(10, workBook.Sheets.Count());
        Assert.Equal("Sheet1", workBook.Sheets.First().Name);
        Assert.Equal("Sheet10", workBook.Sheets.Last().Name);
    }

    [Fact]
    public void WorkBook_SaveProducesNonEmptyBinary()
    {
        var sheet = new WorkSheet("EmptySheet");
        var workBook = new WorkBook("Test", [sheet]);

        var bytes = SheetConverter.ToBinaryExcelFile(workBook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void WorkBook_AllowsDuplicateSheetNames()
    {
        var sheets = new[] { new WorkSheet("Data"), new WorkSheet("Data") };

        var workBook = new WorkBook("Test", sheets);

        Assert.Equal(2, workBook.Sheets.Count());
        Assert.All(workBook.Sheets, sheet => Assert.Equal("Data", sheet.Name));
    }

    [Fact]
    public void WorkBook_SaveAndLoad_PreservesSheetCount()
    {
        var sheets = Enumerable.Range(0, 3).Select(i => new WorkSheet($"Sheet{i}")).ToList();
        var workBook = new WorkBook("RoundTrip", sheets);

        var bytes = SheetConverter.ToBinaryExcelFile(workBook);
        using var stream = new MemoryStream(bytes);
        var loaded = WorkBookReader.LoadFromStream(stream);

        Assert.Equal(sheets.Count, loaded.Sheets.Count());
    }

    [Fact]
    public void WorkSheet_AllowsEmptyName()
    {
        var sheet = new WorkSheet(string.Empty);

        Assert.Equal(string.Empty, sheet.Name);
    }

    [Fact]
    public void WorkSheet_AllowsLongNames()
    {
        var longName = new string('X', 64);

        var sheet = new WorkSheet(longName);

        Assert.Equal(longName, sheet.Name);
    }

    [Fact]
    public void WorkBook_CanSerializeSheetsWithoutCells()
    {
        var sheets = new[] { new WorkSheet("A"), new WorkSheet("B") };
        var workBook = new WorkBook("Test", sheets);

        var bytes = SheetConverter.ToBinaryExcelFile(workBook);

        Assert.NotEmpty(bytes);
    }
}
