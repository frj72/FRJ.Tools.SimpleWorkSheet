using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class WorkSheetBuilderTests
{
    [Fact]
    public void WorkSheet_AddCell_WithCellBuilderAction_AddsCell()
    {
        var sheet = new WorkSheet("Test");
        var position = new CellPosition(1, 1);

        var cell = sheet.AddCell(position, builder => builder
            .WithValue("TestValue")
            .WithColor("FF0000"));

        Assert.Equal("TestValue", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("FF0000", cell.Color);
    }

    [Fact]
    public void WorkSheet_AddCell_WithValueAndBuilder_AddsCell()
    {
        var sheet = new WorkSheet("Test");
        var position = new CellPosition(1, 1);

        var cell = sheet.AddCell(position, "TestValue", builder => builder
            .WithColor("00FF00")
            .WithFont(font => font.Bold()));

        Assert.Equal("TestValue", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("00FF00", cell.Color);
        Assert.True(cell.Font?.Bold);
    }

    [Fact]
    public void WorkSheet_AddCell_WithValueOnly_AddsCell()
    {
        var sheet = new WorkSheet("Test");
        var position = new CellPosition(1, 1);

        var cell = sheet.AddCell(position, "TestValue");

        Assert.Equal("TestValue", sheet.GetValue(1, 1)?.AsString());
    }

    [Fact]
    public void WorkSheet_AddCell_WithCoordinates_AddsCell()
    {
        var sheet = new WorkSheet("Test");

        var cell = sheet.AddCell(1, 1, builder => builder
            .WithValue("TestValue")
            .WithColor("0000FF"));

        Assert.Equal("TestValue", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("0000FF", cell.Color);
    }

    [Fact]
    public void WorkSheet_AddCell_WithCoordinatesAndValue_AddsCell()
    {
        var sheet = new WorkSheet("Test");

        var cell = sheet.AddCell(2, 2, "TestValue", builder => builder
            .WithFont(font => font.WithSize(14).Italic()));

        Assert.Equal("TestValue", sheet.GetValue(2, 2)?.AsString());
        Assert.Equal(14, cell.Font?.Size);
        Assert.True(cell.Font?.Italic);
    }

    [Fact]
    public void WorkSheet_AddStyledCell_WithPosition_AppliesStyle()
    {
        var sheet = new WorkSheet("Test");
        var style = CellStyle.Create("EFEFEF", CellFont.Create(12, "Arial", "000000"), null, null);
        var position = new CellPosition(1, 1);

        var cell = sheet.AddStyledCell(position, "TestValue", style);

        Assert.Equal("TestValue", cell.Value.AsString());
        Assert.Equal(style, cell.Style);
    }

    [Fact]
    public void WorkSheet_AddStyledCell_WithCoordinates_AppliesStyle()
    {
        var sheet = new WorkSheet("Test");
        var style = CellStyle.Create("FFFFFF", CellFont.Create(14, bold: true), null, null);

        var cell = sheet.AddStyledCell(1, 1, "TestValue", style);

        Assert.Equal("TestValue", cell.Value.AsString());
        Assert.Equal(style, cell.Style);
        Assert.True(cell.Font?.Bold);
    }

    [Fact]
    public void WorkSheet_AddRow_AddsMultipleCells()
    {
        var sheet = new WorkSheet("Test");
        var values = new[] { "A", "B", "C" }.Select(s => new CellValue(s));

        var cells = sheet.AddRow(1, 0, values).ToList();

        Assert.Equal(3, cells.Count);
        Assert.Equal("A", sheet.GetValue(0, 1)?.AsString());
        Assert.Equal("B", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("C", sheet.GetValue(2, 1)?.AsString());
    }

    [Fact]
    public void WorkSheet_AddRow_WithConfiguration_AppliesStyleToAll()
    {
        var sheet = new WorkSheet("Test");
        var values = new[] { "A", "B", "C" }.Select(s => new CellValue(s));

        var cells = sheet.AddRow(1, 0, values, builder => builder
            .WithColor("EFEFEF")
            .WithFont(font => font.Bold())).ToList();

        Assert.All(cells, cell =>
        {
            Assert.Equal("EFEFEF", cell.Color);
            Assert.True(cell.Font?.Bold);
        });
    }

    [Fact]
    public void WorkSheet_AddColumn_AddsMultipleCells()
    {
        var sheet = new WorkSheet("Test");
        var values = new[] { "A", "B", "C" }.Select(s => new CellValue(s));

        var cells = sheet.AddColumn(1, 0, values).ToList();

        Assert.Equal(3, cells.Count);
        Assert.Equal("A", sheet.GetValue(1, 0)?.AsString());
        Assert.Equal("B", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("C", sheet.GetValue(1, 2)?.AsString());
    }

    [Fact]
    public void WorkSheet_AddColumn_WithConfiguration_AppliesStyleToAll()
    {
        var sheet = new WorkSheet("Test");
        var values = new[] { 1, 2, 3 }.Select(i => new CellValue(i));

        var cells = sheet.AddColumn(1, 0, values, builder => builder
            .WithFont(font => font.WithSize(16))).ToList();

        Assert.All(cells, cell => Assert.Equal(16, cell.Font?.Size));
    }

    [Fact]
    public void WorkSheet_UpdateCell_ModifiesExistingCell()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(1, 1, "Original");

        var updated = sheet.UpdateCell(1, 1, builder => builder
            .WithValue("Updated")
            .WithColor("FF0000"));

        Assert.Equal("Updated", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("FF0000", updated.Color);
    }

    [Fact]
    public void WorkSheet_UpdateCell_CreatesNewCellIfNotExists()
    {
        var sheet = new WorkSheet("Test");

        var cell = sheet.UpdateCell(1, 1, builder => builder
            .WithValue("NewCell")
            .WithColor("00FF00"));

        Assert.Equal("NewCell", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("00FF00", cell.Color);
    }

    [Fact]
    public void WorkSheet_AddCell_ComplexStyling_WorksCorrectly()
    {
        var sheet = new WorkSheet("Test");

        var cell = sheet.AddCell(1, 1, 123.45m, builder => builder
            .WithStyle(style => style
                .WithFillColor("E0E0E0")
                .WithFont(font => font
                    .WithSize(14)
                    .WithName("Calibri")
                    .Bold()
                    .Italic())
                .WithFormatCode("0.00")));

        Assert.Equal(123.45m, cell.Value.AsDecimal());
        Assert.Equal("E0E0E0", cell.Color);
        Assert.Equal(14, cell.Font?.Size);
        Assert.Equal("Calibri", cell.Font?.Name);
        Assert.True(cell.Font?.Bold);
        Assert.True(cell.Font?.Italic);
        Assert.Equal("0.00", cell.FormatCode);
    }

    [Fact]
    public void WorkSheet_AddCell_WithMetadata_StoresMetadata()
    {
        var sheet = new WorkSheet("Test");

        var cell = sheet.AddCell(1, 1, "Imported", builder => builder
            .WithMetadata(meta => meta
                .WithSource("csv")
                .WithOriginalValue("raw_value")
                .AddCustomData("row", 10)));

        Assert.NotNull(cell.Metadata);
        Assert.Equal("csv", cell.Metadata.Source);
        Assert.Equal("raw_value", cell.Metadata.OriginalValue);
        Assert.Equal(10, cell.Metadata.CustomData?["row"]);
    }

    [Fact]
    public void WorkSheet_AddCell_InvalidColor_ThrowsException()
    {
        var sheet = new WorkSheet("Test");

        Assert.Throws<ArgumentException>(() => 
            sheet.AddCell(1, 1, "Test", builder => builder.WithColor("invalidColor")));
    }

    [Fact]
    public void WorkSheet_NewAddCellAPI_Works()
    {
        var sheet = new WorkSheet("Test");
        var position = new CellPosition(1, 1);

        var cell = sheet.AddCell(position, "TestValue", builder => builder.WithColor("FF0000"));

        Assert.Equal("TestValue", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal("FF0000", cell.Color);
    }
}
