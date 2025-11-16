using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellMergingTests
{
    [Fact]
    public void SheetConverter_WritesMergedCells()
    {
        var sheet = new WorkSheet("Merged");
        sheet.AddCell(new(0, 0), "Header", cell => cell.WithFont(f => f.Bold()));
        sheet.MergeCells(0, 0, 2, 0);

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart?.WorksheetParts.First();
        var mergeCells = worksheetPart?.Worksheet.Elements<MergeCells>().FirstOrDefault();

        Assert.NotNull(mergeCells);
        var mergeCell = mergeCells.Elements<MergeCell>().FirstOrDefault();
        Assert.Equal("A1:C1", mergeCell?.Reference?.Value);
    }

    [Fact]
    public void WorkBookReader_LoadsMergedCells()
    {
        var sheet = new WorkSheet("Merged");
        sheet.AddCell(new(0, 0), "Header");
        sheet.MergeCells(0, 0, 3, 1);

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        using var stream = new MemoryStream(bytes);
        var loaded = WorkBookReader.LoadFromStream(stream);
        var loadedSheet = loaded.Sheets.First();

        Assert.Single(loadedSheet.MergedCells);
        var range = loadedSheet.MergedCells[0];
        Assert.Equal(new(0, 0), range.From);
        Assert.Equal(new(3, 1), range.To);
    }

    [Fact]
    public void MergeCells_WithSingleCellRange_ThrowsArgumentException()
    {
        var sheet = new WorkSheet("Test");

        var exception = Assert.Throws<ArgumentException>(() => sheet.MergeCells(1, 1, 1, 1));
        Assert.Equal("range", exception.ParamName);
    }

    [Fact]
    public void MergeCells_WithExtremeCoordinates_AddsRange()
    {
        var sheet = new WorkSheet("Test");

        sheet.MergeCells(0, 0, 16383, 1048575);

        Assert.Single(sheet.MergedCells);
        var range = sheet.MergedCells[0];
        Assert.Equal(new(0, 0), range.From);
        Assert.Equal(new(16383, 1048575), range.To);
    }

    [Fact]
    public void MergeCells_PreservesExistingTopLeftValue()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Header");

        sheet.MergeCells(0, 0, 2, 1);

        var cell = sheet.Cells.Cells[new(0, 0)];
        Assert.Equal("Header", cell.Value.Value.AsT2);
    }

    [Fact]
    public void MergeCells_CreatesTopLeftCellWhenMissing()
    {
        var sheet = new WorkSheet("Test");

        sheet.MergeCells(3, 4, 5, 6);

        var position = new CellPosition(3, 4);
        Assert.True(sheet.Cells.Cells.ContainsKey(position));
        var cell = sheet.Cells.Cells[position];
        Assert.True(cell.Value.Value.IsT2);
        Assert.Equal(string.Empty, cell.Value.Value.AsT2);
    }
}
