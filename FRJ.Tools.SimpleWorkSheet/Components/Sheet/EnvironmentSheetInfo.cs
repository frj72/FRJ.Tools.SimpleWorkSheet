using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class EnvironmentSheetInfo
{
    private static readonly DateTime ExampleDate = new(2027, 12, 12, 23, 12, 38, DateTimeKind.Unspecified);
    public static double GetWidth(string fontName, int fontSize, string text, bool bold = false, bool italic = false)
    {
        var font = CellFont.Create(fontSize, fontName, Colors.Black, bold, italic);
        var style = CellStyle.Create(Colors.White, font, WorkSheetDefaults.CellBorders);
        var cell = new Cell(text, style, null);
        
        return new[] { cell }.EstimateMaxWidth();
    }

    public static WorkBook CreateSimpleCalibrationWorkBook(double? calibrationFactor = null)
    {
        var genericTable = GenericTable.Create("Factor", "Integer", "Float", "Text One", "Text Two", "date");
        genericTable.AddRow(calibrationFactor ?? 1.0, 1000, 1000.00001m, "This is a simple calibration test", "All cell values should be fitted", ExampleDate);

        var builder = WorkbookBuilder
            .FromGenericTable(genericTable)
            .WithDataSheetName("Calibration Sheet")
            .WithWorkbookName("Calibration Workbook")
            .WithHeaderStyle(style => style
                .WithFont(font => font.Bold().WithSize(12))
                .WithFillColor(Colors.Platinum)
                );
        if (calibrationFactor is null)
            builder.AutoFitAllColumns();
        else
            builder.AutoFitAllColumns(calibrationFactor.Value);
        return builder.Build();
    }
}
