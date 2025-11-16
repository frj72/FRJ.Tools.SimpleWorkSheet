using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class NamedRangeTests
{
    [Fact]
    public void NamedRange_Constructor_SetsProperties()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);
        var namedRange = new NamedRange("TestRange", "Sheet1", range);

        Assert.Equal("TestRange", namedRange.Name);
        Assert.Equal("Sheet1", namedRange.SheetName);
        Assert.Equal(range, namedRange.Range);
    }

    [Fact]
    public void NamedRange_Constructor_ThrowsOnNullName()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);

        Assert.Throws<ArgumentException>(() => new NamedRange("", "Sheet1", range));
        Assert.Throws<ArgumentException>(() => new NamedRange(null, "Sheet1", range));
    }

    [Fact]
    public void NamedRange_Constructor_ThrowsOnNullSheetName()
    {
        var range = CellRange.FromBounds(0, 0, 5, 10);

        Assert.Throws<ArgumentException>(() => new NamedRange("TestRange", "", range));
        Assert.Throws<ArgumentException>(() => new NamedRange("TestRange", null, range));
    }

    [Fact]
    public void ToFormulaReference_GeneratesCorrectFormat()
    {
        var range = CellRange.FromBounds(0, 0, 2, 5);
        var namedRange = new NamedRange("TestRange", "Sheet1", range);

        var formula = namedRange.ToFormulaReference();

        Assert.Equal("'Sheet1'!$A$1:$C$6", formula);
    }

    [Fact]
    public void ToFormulaReference_HandlesLargeColumns()
    {
        var range = CellRange.FromBounds(25, 0, 27, 5);
        var namedRange = new NamedRange("TestRange", "Sheet1", range);

        var formula = namedRange.ToFormulaReference();

        Assert.Equal("'Sheet1'!$Z$1:$AB$6", formula);
    }

    [Fact]
    public void WorkBook_AddNamedRange_AddsToCollection()
    {
        var sheet = new WorkSheet("Sheet1");
        var workbook = new WorkBook("Test", [sheet]);
        var range = CellRange.FromBounds(0, 0, 5, 10);

        workbook.AddNamedRange("TestRange", "Sheet1", range);

        Assert.Single(workbook.NamedRanges);
        Assert.Equal("TestRange", workbook.NamedRanges[0].Name);
    }

    [Fact]
    public void WorkBook_AddNamedRange_WithCoordinates_AddsToCollection()
    {
        var sheet = new WorkSheet("Sheet1");
        var workbook = new WorkBook("Test", [sheet]);

        workbook.AddNamedRange("TestRange", "Sheet1", 0, 0, 5, 10);

        Assert.Single(workbook.NamedRanges);
        var namedRange = workbook.NamedRanges[0];
        Assert.Equal("TestRange", namedRange.Name);
        Assert.Equal(new(0, 0), namedRange.Range.From);
        Assert.Equal(new(5, 10), namedRange.Range.To);
    }

    [Fact]
    public void WorkBook_AddNamedRange_ThrowsOnDuplicateName()
    {
        var sheet = new WorkSheet("Sheet1");
        var workbook = new WorkBook("Test", [sheet]);
        var range1 = CellRange.FromBounds(0, 0, 5, 10);
        var range2 = CellRange.FromBounds(6, 0, 10, 10);

        workbook.AddNamedRange("TestRange", "Sheet1", range1);

        Assert.Throws<ArgumentException>(() =>
            workbook.AddNamedRange("TestRange", "Sheet1", range2));
    }

    [Fact]
    public void WorkBook_AddNamedRange_IsCaseInsensitive()
    {
        var sheet = new WorkSheet("Sheet1");
        var workbook = new WorkBook("Test", [sheet]);
        var range1 = CellRange.FromBounds(0, 0, 5, 10);
        var range2 = CellRange.FromBounds(6, 0, 10, 10);

        workbook.AddNamedRange("TestRange", "Sheet1", range1);

        Assert.Throws<ArgumentException>(() =>
            workbook.AddNamedRange("testrange", "Sheet1", range2));
    }

    [Fact]
    public void NamedRange_SingleCell_GeneratesCorrectFormat()
    {
        var range = CellRange.FromBounds(3, 5, 3, 5);
        var namedRange = new NamedRange("SingleCell", "Sheet1", range);

        var formula = namedRange.ToFormulaReference();

        Assert.Equal("'Sheet1'!$D$6:$D$6", formula);
    }
}
