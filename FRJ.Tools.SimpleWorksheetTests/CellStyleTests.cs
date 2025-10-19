using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellStyleTests
{
    [Fact]
    public void CellStyle_Create_WithAllParameters_SetsAllProperties()
    {
        var fillColor = "FF0000";
        var font = CellFont.Create(14, "Arial", "000000", true, false, false, false);
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        var formatCode = "0.00";

        var style = CellStyle.Create(fillColor, font, borders, formatCode);

        Assert.Equal(fillColor, style.FillColor);
        Assert.Equal(font, style.Font);
        Assert.Equal(borders, style.Borders);
        Assert.Equal(formatCode, style.FormatCode);
    }

    [Fact]
    public void CellStyle_Create_WithNullParameters_SetsNullProperties()
    {
        var style = CellStyle.Create(null, null, null, null);

        Assert.Null(style.FillColor);
        Assert.Null(style.Font);
        Assert.Null(style.Borders);
        Assert.Null(style.FormatCode);
    }

    [Fact]
    public void CellStyle_Create_WithNoParameters_SetsNullProperties()
    {
        var style = CellStyle.Create();

        Assert.Null(style.FillColor);
        Assert.Null(style.Font);
        Assert.Null(style.Borders);
        Assert.Null(style.FormatCode);
    }

    [Fact]
    public void CellStyle_RecordEquality_WithSameValues_AreEqual()
    {
        var fillColor = "FFFFFF";
        var font = CellFont.Create(12, "Arial", "000000");
        var style1 = CellStyle.Create(fillColor, font, null, null);
        var style2 = CellStyle.Create(fillColor, font, null, null);

        Assert.Equal(style1, style2);
    }

    [Fact]
    public void CellStyle_RecordEquality_WithDifferentValues_AreNotEqual()
    {
        var style1 = CellStyle.Create("FFFFFF", null, null, null);
        var style2 = CellStyle.Create("000000", null, null, null);

        Assert.NotEqual(style1, style2);
    }

    [Fact]
    public void CellStyle_WithExpression_UpdatesProperty()
    {
        var originalStyle = CellStyle.Create("FFFFFF", null, null, null);
        var newFont = CellFont.Create(14, "Arial", "000000");
        
        var updatedStyle = originalStyle with { Font = newFont };

        Assert.Equal("FFFFFF", updatedStyle.FillColor);
        Assert.Equal(newFont, updatedStyle.Font);
    }
}
