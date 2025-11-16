using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class WorkSheetTests
{
    [Fact]
    public void AddCell_ValidParameters_CellAdded()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(1, 1);
        var value = new CellValue("TestValue");

        sheet.AddCell(position, value);

        Assert.Equal(value, sheet.GetValue(1, 1));
    }

    [Fact]
    public void AddCell_InvalidColor_ThrowsArgumentException()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(1, 1);
        var value = new CellValue("TestValue");

        Assert.Throws<ArgumentException>(() => sheet.AddCell(position, value, builder => builder.WithColor("invalidColor")));
    }

    [Fact]
    public void SetValue_ExistingCell_UpdatesValue()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(1, 1);
        var value = new CellValue("TestValue");

        sheet.AddCell(position, value);
        var newValue = new CellValue("NewValue");
        sheet.SetValue(1, 1, newValue);

        Assert.Equal(newValue, sheet.GetValue(1, 1));
    }

    [Fact]
    public void SetValue_NewCell_AddsCell()
    {
        var sheet = new WorkSheet("TestSheet");
        var value = new CellValue("TestValue");

        sheet.SetValue(1, 1, value);

        Assert.Equal(value, sheet.GetValue(1, 1));
    }

    [Fact]
    public void SetFont_ValidFont_UpdatesFont()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(1, 1);
        var font = CellFont.Create(12, "Arial", "000000");

        sheet.AddCell(position, new CellValue("TestValue"));
        sheet.SetFont(1, 1, font);

        Assert.Equal(font, sheet.Cells.Cells[position].Font);
    }

    [Fact]
    public void SetColor_ValidColor_UpdatesColor()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(1, 1);
        const string color = "FFFFFF";

        sheet.AddCell(position, new CellValue("TestValue"));
        sheet.SetColor(1, 1, color);

        Assert.Equal(color, sheet.Cells.Cells[position].Color);
    }

    [Fact]
    public void SetBorders_ValidBorders_UpdatesBorders()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(1, 1);
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));

        sheet.AddCell(position, new CellValue("TestValue"));
        sheet.SetBorders(1, 1, borders);

        Assert.Equal(borders, sheet.Cells.Cells[position].Borders);
    }
    
    [Fact]
    public void SetFont_InvalidFontColor_ThrowsArgumentException()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(1, 1);
        var font = CellFont.Create(12, "Arial", "invalidColor");
 
        sheet.AddCell(position, new CellValue("TestValue"));
 
        Assert.Throws<ArgumentException>(() => sheet.SetFont(1, 1, font));
    }

    [Fact]
    public void MergeCells_AddsRangeAndEnsuresTopLeftCellExists()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.MergeCells(0, 0, 2, 1);
 
        Assert.Single(sheet.MergedCells);
        Assert.True(sheet.Cells.Cells.ContainsKey(new(0, 0)));
    }

    [Fact]
    public void MergeCells_OverlappingRange_Throws()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.MergeCells(0, 0, 1, 0);
 
        Assert.Throws<ArgumentException>(() => sheet.MergeCells(1, 0, 2, 0));
    }

    [Fact]
    public void Charts_DefaultValue_IsEmpty()
    {
        var sheet = new WorkSheet("TestSheet");

        Assert.Empty(sheet.Charts);
    }

    [Fact]
    public void AddChart_ValidChart_AddsToCollection()
    {
        var sheet = new WorkSheet("TestSheet");
        var chart = BarChart.Create()
            .WithTitle("Test Chart")
            .WithPosition(5, 0, 10, 15);

        sheet.AddChart(chart);

        Assert.Single(sheet.Charts);
        Assert.Same(chart, sheet.Charts[0]);
    }

    [Fact]
    public void AddChart_MultipleCharts_AddsAll()
    {
        var sheet = new WorkSheet("TestSheet");
        var chart1 = BarChart.Create()
            .WithTitle("Chart 1")
            .WithPosition(5, 0, 10, 15);
        var chart2 = BarChart.Create()
            .WithTitle("Chart 2")
            .WithPosition(12, 0, 17, 15);

        sheet.AddChart(chart1);
        sheet.AddChart(chart2);

        Assert.Equal(2, sheet.Charts.Count);
        Assert.Same(chart1, sheet.Charts[0]);
        Assert.Same(chart2, sheet.Charts[1]);
    }

    [Fact]
    public void AddChart_ChartWithoutPosition_ThrowsArgumentException()
    {
        var sheet = new WorkSheet("TestSheet");
        var chart = BarChart.Create()
            .WithTitle("Test Chart");

        var ex = Assert.Throws<ArgumentException>(() => sheet.AddChart(chart));
        Assert.Contains("position", ex.Message.ToLower());
    }

    [Fact]
    public void AddChart_ChartsAreReadOnly_CannotModifyDirectly()
    {
        var sheet = new WorkSheet("TestSheet");
        var chart = BarChart.Create()
            .WithTitle("Test Chart")
            .WithPosition(5, 0, 10, 15);

        sheet.AddChart(chart);

        Assert.IsType<IReadOnlyList<Chart>>(sheet.Charts, exactMatch: false);
    }
}

