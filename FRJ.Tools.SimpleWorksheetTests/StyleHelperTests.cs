using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class StyleHelperTests
{
    [Fact]
    public void CollectStyles_ReusesFillAndStyleForIdenticalCells()
    {
        const string fillColor = "FF112233";

        var sheet = new WorkSheet("Sheet1");
        var firstCell = sheet.AddCell(new(0, 0), "First", builder => builder.WithColor(fillColor));
        var secondCell = sheet.AddCell(new(1, 0), "Second", builder => builder.WithColor(fillColor));
        var workBook = new WorkBook("Test", [sheet]);
        var styleHelper = new StyleHelper();

        styleHelper.CollectStyles(workBook);
        var stylesheet = styleHelper.GenerateStylesheet();

        var fills = Assert.IsType<Fills>(stylesheet.Fills);
        Assert.Equal<uint>(3, fills.Count?.Value ?? 0);
        var firstIndex = styleHelper.GetStyleIndex(firstCell);
        var secondIndex = styleHelper.GetStyleIndex(secondCell);
        Assert.Equal(firstIndex, secondIndex);
    }

    [Fact]
    public void CollectStyles_ForDateCell_AppliesDateNumberFormat()
    {
        var sheet = new WorkSheet("Sheet1");
        var dateCell = sheet.AddCell(new(0, 0), new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Utc), null);
        var workBook = new WorkBook("Test", [sheet]);
        var styleHelper = new StyleHelper();

        styleHelper.CollectStyles(workBook);
        var styleIndex = styleHelper.GetStyleIndex(dateCell);
        var stylesheet = styleHelper.GenerateStylesheet();
        var cellFormats = Assert.IsType<CellFormats>(stylesheet.CellFormats);
        var cellFormat = Assert.IsType<CellFormat>(cellFormats.ChildElements[(int)styleIndex]);

        Assert.Equal<uint>(164, cellFormat.NumberFormatId?.Value ?? 0);
        Assert.True(cellFormat.ApplyNumberFormat?.Value ?? false);
    }

    [Fact]
    public void CollectStyles_WithAlignmentInformation_CreatesAlignmentElement()
    {
        var sheet = new WorkSheet("Sheet1");
        var alignedCell = sheet.AddCell(new(0, 0), "Aligned", builder =>
            builder.WithStyle(style => style
                .WithHorizontalAlignment(HorizontalAlignment.Center)
                .WithVerticalAlignment(VerticalAlignment.Bottom)
                .WithTextRotation(30)
                .WithWrapText(true)));
        var workBook = new WorkBook("Test", [sheet]);
        var styleHelper = new StyleHelper();

        styleHelper.CollectStyles(workBook);
        var styleIndex = styleHelper.GetStyleIndex(alignedCell);
        var stylesheet = styleHelper.GenerateStylesheet();
        var cellFormats = Assert.IsType<CellFormats>(stylesheet.CellFormats);
        var cellFormat = Assert.IsType<CellFormat>(cellFormats.ChildElements[(int)styleIndex]);

        Assert.True(cellFormat.ApplyAlignment?.Value ?? false);
        var alignment = Assert.IsType<Alignment>(cellFormat.Alignment);
        Assert.Equal(HorizontalAlignmentValues.Center, alignment.Horizontal?.Value);
        Assert.Equal(VerticalAlignmentValues.Bottom, alignment.Vertical?.Value);
        Assert.Equal<uint>(30, alignment.TextRotation?.Value ?? 0);
        Assert.True(alignment.WrapText ?? false);
    }
}
