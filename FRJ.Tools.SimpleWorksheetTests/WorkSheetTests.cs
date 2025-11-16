using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

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

    [Fact]
    public void AddChart_WithDataAndChart_GeneratesValidExcelFile()
    {
        var sheet = new WorkSheet("Sales");

        sheet.AddCell(new(0, 0), new CellValue("Region"));
        sheet.AddCell(new(1, 0), new CellValue("Sales"));
        sheet.AddCell(new(0, 1), new CellValue("North"));
        sheet.AddCell(new(1, 1), new CellValue(100));
        sheet.AddCell(new(0, 2), new CellValue("South"));
        sheet.AddCell(new(1, 2), new CellValue(150));
        sheet.AddCell(new(0, 3), new CellValue("East"));
        sheet.AddCell(new(1, 3), new CellValue(120));

        var categoriesRange = CellRange.FromBounds(0, 1, 0, 3);
        var valuesRange = CellRange.FromBounds(1, 1, 1, 3);

        var chart = BarChart.Create()
            .WithTitle("Sales by Region")
            .WithDataRange(categoriesRange, valuesRange)
            .WithPosition(3, 0, 10, 15)
            .WithOrientation(BarChartOrientation.Vertical);

        sheet.AddChart(chart);

        var excelBytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.NotNull(excelBytes);
        Assert.True(excelBytes.Length > 0);
        Assert.True(excelBytes.Length > 3000, $"File size was {excelBytes.Length} bytes");
    }

    [Fact]
    public void AddChart_ChartOnlySheet_NoCells_Works()
    {
        var chartSheet = new WorkSheet("Dashboard");

        var chart = BarChart.Create()
            .WithTitle("Chart Only")
            .WithDataRange(CellRange.FromBounds(0, 1, 0, 3), CellRange.FromBounds(1, 1, 1, 3))
            .WithPosition(0, 0, 8, 15);

        chartSheet.AddChart(chart);

        Assert.Empty(chartSheet.Cells.Cells);
        Assert.Single(chartSheet.Charts);
    }

    [Fact]
    public void AddChart_MultipleSheets_ChartsReferenceCrossSheetsCorrectly()
    {
        var dataSheet = new WorkSheet("Data");
        dataSheet.AddCell(new(0, 0), new CellValue("A"));
        dataSheet.AddCell(new(1, 0), new CellValue(100));
        dataSheet.AddCell(new(0, 1), new CellValue("B"));
        dataSheet.AddCell(new(1, 1), new CellValue(200));

        var chartSheet = new WorkSheet("Charts");
        var chart = BarChart.Create()
            .WithTitle("Cross Sheet Chart")
            .WithDataRange(CellRange.FromBounds(0, 0, 0, 1), CellRange.FromBounds(1, 0, 1, 1))
            .WithPosition(0, 0, 8, 15);

        chartSheet.AddChart(chart);

        var workbook = new WorkBook("Test", [dataSheet, chartSheet]);
        var excelBytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotNull(excelBytes);
        Assert.True(excelBytes.Length > 0);
    }

    [Fact]
    public void AddCell_ToSamePositionTwice_SecondOverwritesFirst()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(0, 0);

        sheet.AddCell(position, "First");
        sheet.AddCell(position, "Second");

        var value = sheet.GetValue(0, 0);
        Assert.Equal("Second", value?.Value.AsT2);
    }

    [Fact]
    public void AddCell_ThenSetValue_ValueUpdates()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(0, 0);

        sheet.AddCell(position, "Original");
        sheet.SetValue(0, 0, "Updated");

        var value = sheet.GetValue(0, 0);
        Assert.Equal("Updated", value?.Value.AsT2);
    }

    [Fact]
    public void AddCell_ThenSetFont_FontUpdatesValuePreserved()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(0, 0);
        var newFont = CellFont.Create(16, "Arial", "FF0000");

        sheet.AddCell(position, "TestValue");
        sheet.SetFont(0, 0, newFont);

        var value = sheet.GetValue(0, 0);
        var cell = sheet.Cells.Cells[position];
        
        Assert.Equal("TestValue", value?.Value.AsT2);
        Assert.Equal(newFont, cell.Font);
    }

    [Fact]
    public void AddCell_ThenSetColor_ColorUpdatesValuePreserved()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(0, 0);
        const string newColor = "FF0000";

        sheet.AddCell(position, "TestValue");
        sheet.SetColor(0, 0, newColor);

        var value = sheet.GetValue(0, 0);
        var cell = sheet.Cells.Cells[position];
        
        Assert.Equal("TestValue", value?.Value.AsT2);
        Assert.Equal(newColor, cell.Color);
    }

    [Fact]
    public void AddCell_ThenSetBorders_BordersUpdateValuePreserved()
    {
        var sheet = new WorkSheet("TestSheet");
        var position = new CellPosition(0, 0);
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));

        sheet.AddCell(position, "TestValue");
        sheet.SetBorders(0, 0, borders);

        var value = sheet.GetValue(0, 0);
        var cell = sheet.Cells.Cells[position];
        
        Assert.Equal("TestValue", value?.Value.AsT2);
        Assert.Equal(borders, cell.Borders);
    }

    [Fact]
    public void SetColumnWidth_WithZero_StoresValue()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.SetColumnWith(0, 0.0);

        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.Equal(0.0, sheet.ExplicitColumnWidths[0].AsT0);
    }

    [Fact]
    public void SetColumnWidth_WithNegativeValue_StoresValue()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.SetColumnWith(0, -5.0);

        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.Equal(-5.0, sheet.ExplicitColumnWidths[0].AsT0);
    }

    [Fact]
    public void SetColumnWidth_WithMinimum_StoresValue()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.SetColumnWith(0, 1.0);

        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.Equal(1.0, sheet.ExplicitColumnWidths[0].AsT0);
    }

    [Fact]
    public void SetColumnWidth_WithExcelMaximum_StoresValue()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.SetColumnWith(0, 255.0);

        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.Equal(255.0, sheet.ExplicitColumnWidths[0].AsT0);
    }

    [Fact]
    public void SetColumnWidth_WithBeyondExcelLimit_StoresValue()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.SetColumnWith(0, 1000.0);

        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.Equal(1000.0, sheet.ExplicitColumnWidths[0].AsT0);
    }

    [Fact]
    public void SetColumnWidth_WithAutoEnum_StoresValue()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.SetColumnWith(0, CellWidth.AutoExpand);

        Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
        Assert.True(sheet.ExplicitColumnWidths[0].IsT1);
        Assert.Equal(CellWidth.AutoExpand, sheet.ExplicitColumnWidths[0].AsT1);
    }
}

