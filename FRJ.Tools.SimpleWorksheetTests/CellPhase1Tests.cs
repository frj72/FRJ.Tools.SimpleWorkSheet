using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellPhase1Tests
{
    [Fact]
    public void Cell_ConstructorWithStyle_SetsAllProperties()
    {
        var value = new CellValue("TestValue");
        var style = CellStyle.Create("FFFFFF", CellFont.Create(12, "Arial", "000000"), null, "0.00");
        var metadata = CellMetadata.Create("csv", DateTime.UtcNow, "raw_value");

        var cell = new Cell(value, style, metadata);

        Assert.Equal(value, cell.Value);
        Assert.Equal(style, cell.Style);
        Assert.Equal(metadata, cell.Metadata);
    }

    [Fact]
    public void Cell_ConstructorWithStyle_WithoutMetadata_SetsNullMetadata()
    {
        var value = new CellValue("TestValue");
        var style = CellStyle.Create("FFFFFF");

        var cell = new Cell(value, style, null);

        Assert.Equal(value, cell.Value);
        Assert.Equal(style, cell.Style);
        Assert.Null(cell.Metadata);
    }

    [Fact]
    public void Cell_LegacyConstructor_CreatesStyleInternally()
    {
        var value = new CellValue("TestValue");
        const string color = "FFFFFF";
        var font = CellFont.Create(12, "Arial", "000000");
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        const string formatCode = "0.00";

        var cell = new Cell(value, color, font, borders, formatCode);

        Assert.Equal(value, cell.Value);
        Assert.Equal(color, cell.Style.FillColor);
        Assert.Equal(font, cell.Style.Font);
        Assert.Equal(borders, cell.Style.Borders);
        Assert.Equal(formatCode, cell.Style.FormatCode);
    }

    [Fact]
    public void Cell_BackwardCompatibilityProperties_ReturnStyleProperties()
    {
        const string color = "FFFFFF";
        var font = CellFont.Create(12, "Arial", "000000");
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        const string formatCode = "0.00";
        var style = CellStyle.Create(color, font, borders, formatCode);
        var cell = new Cell(new("Test"), style, null);

        Assert.Equal(color, cell.Color);
        Assert.Equal(font, cell.Font);
        Assert.Equal(borders, cell.Borders);
        Assert.Equal(formatCode, cell.FormatCode);
    }

    [Fact]
    public void Cell_WithStyleUpdate_UpdatesStyleProperty()
    {
        var originalStyle = CellStyle.Create("FFFFFF");
        var cell = new Cell(new("Test"), originalStyle, null);
        var newStyle = CellStyle.Create("FF0000");

        var updatedCell = cell with { Style = newStyle };

        Assert.Equal("FF0000", updatedCell.Style.FillColor);
        Assert.Equal("FF0000", updatedCell.Color);
    }

    [Fact]
    public void Cell_WithMetadata_StoresMetadata()
    {
        var style = CellStyle.Create("FFFFFF");
        var metadata = CellMetadata.Create("json", DateTime.UtcNow, "original_value");
        var cell = new Cell(new("Test"), style, metadata);

        Assert.NotNull(cell.Metadata);
        Assert.Equal("json", cell.Metadata.Source);
        Assert.Equal("original_value", cell.Metadata.OriginalValue);
    }

    [Fact]
    public void Cell_Create_WorksWithNewStyleApproach()
    {
        var cell = Cell.Create("TestValue", "FFFFFF", CellFont.Create(12, "Arial", "000000"), null, null);

        Assert.Equal("TestValue", cell.Value.AsString());
        Assert.Equal("FFFFFF", cell.Style.FillColor);
        Assert.NotNull(cell.Style.Font);
        Assert.Equal(12, cell.Style.Font.Size);
    }

    [Fact]
    public void CellExtensions_SetColor_UpdatesStyleFillColor()
    {
        var cell = Cell.Create("Test", null, null, null, null);
        
        var updatedCell = cell.SetColor("FF0000");

        Assert.Equal("FF0000", updatedCell.Style.FillColor);
        Assert.Equal("FF0000", updatedCell.Color);
    }

    [Fact]
    public void CellExtensions_SetFont_UpdatesStyleFont()
    {
        var cell = Cell.Create("Test", null, null, null, null);
        var newFont = CellFont.Create(14, "Arial", "000000", true);
        
        var updatedCell = cell.SetFont(newFont);

        Assert.Equal(newFont, updatedCell.Style.Font);
        Assert.Equal(newFont, updatedCell.Font);
    }

    [Fact]
    public void CellExtensions_SetBorders_UpdatesStyleBorders()
    {
        var cell = Cell.Create("Test", null, null, null, null);
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        var updatedCell = cell.SetBorders(borders);

        Assert.Equal(borders, updatedCell.Style.Borders);
        Assert.Equal(borders, updatedCell.Borders);
    }

    [Fact]
    public void CellExtensions_SetDefaultFormatting_UsesDefaultCellStyle()
    {
        var cell = Cell.Create("Test", "FF0000", CellFont.Create(14, "Arial", "000000", true), null, null);
        
        var updatedCell = cell.SetDefaultFormatting();

        Assert.Equal(WorkSheetDefaults.DefaultCellStyle, updatedCell.Style);
        Assert.Equal(WorkSheetDefaults.FillColor, updatedCell.Color);
        Assert.Equal(WorkSheetDefaults.Font, updatedCell.Font);
    }

    [Fact]
    public void WorkSheetDefaults_DefaultCellStyle_HasCorrectProperties()
    {
        var defaultStyle = WorkSheetDefaults.DefaultCellStyle;

        Assert.Equal(WorkSheetDefaults.FillColor, defaultStyle.FillColor);
        Assert.Equal(WorkSheetDefaults.Font, defaultStyle.Font);
        Assert.Equal(WorkSheetDefaults.CellBorders, defaultStyle.Borders);
    }
}
