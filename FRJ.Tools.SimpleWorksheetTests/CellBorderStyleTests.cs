using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class CellBorderStyleTests
{
    [Fact]
    public void CellBorderStyle_None_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.None);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "None");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_Thin_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.Thin);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "Thin");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_Medium_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.Medium);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "Medium");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_Dashed_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.Dashed);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "Dashed");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_Dotted_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.Dotted);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "Dotted");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_Thick_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.Thick);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "Thick");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_Double_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.Double);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "Double");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_Hair_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.Hair);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "Hair");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_MediumDashed_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.MediumDashed);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "MediumDashed");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_DashDot_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.DashDot);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "DashDot");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_MediumDashDot_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.MediumDashDot);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "MediumDashDot");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_DashDotDot_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.DashDotDot);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "DashDotDot");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_MediumDashDotDot_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.MediumDashDotDot);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "MediumDashDotDot");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellBorderStyle_SlantDashDot_CreatesValidExcelFile()
    {
        var border = CellBorder.Create(Colors.Black, CellBorderStyle.SlantDashDot);
        var borders = CellBorders.Create(border, border, border, border);
        var sheet = CreateSheetWithBorder(borders, "SlantDashDot");

        var bytes = SheetConverter.ToBinaryExcelFile(sheet);

        Assert.True(bytes.Length > 1000);
    }

    private static WorkSheet CreateSheetWithBorder(CellBorders borders, string label)
    {
        var sheet = new WorkSheet("BorderTest");
        sheet.AddCell(new(0, 0), label, cell => cell
            .WithStyle(style => style.WithBorders(borders)));
        return sheet;
    }
}
