using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class AutoFitColumnsBaselineExample : IExample
{
    public string Name => "Auto-Fit Columns with Baseline";
    public string Description => "Demonstrate baseline parameter for adding fixed offset to column widths";

    public int ExampleNumber { get; }

    public AutoFitColumnsBaselineExample(int exampleNumber) => ExampleNumber = exampleNumber;

    private static readonly string[] Headers = ["Product", "Price", "Stock"];
    private static readonly object[][] Data =
    [
        ["Laptop", 1299.99m, 15],
        ["Mouse", 24.99m, 150],
        ["Keyboard", 89.99m, 75],
        ["Monitor", 349.99m, 30],
        ["USB Cable", 9.99m, 200]
    ];

    public void Run()
    {
        var sheet1 = CreateSheetWithBaseline("No Baseline", 1.0, 0.0, 0);
        var sheet2 = CreateSheetWithBaseline("Baseline +5", 1.0, 5.0, 8);
        var sheet3 = CreateSheetWithBaseline("Baseline +10", 1.0, 10.0, 16);
        var sheet4 = CreateSheetWithBaseline("Baseline -3", 1.0, -3.0, 24);
        var sheet5 = CreateSheetWithBaseline("Calibration 1.2 Baseline +8", 1.2, 8.0, 32);

        var workbook = new WorkBook("AutoFit Baseline Demo", [sheet1, sheet2, sheet3, sheet4, sheet5]);

        ExampleRunner.SaveWorkBook(workbook, $"{ExampleNumber:000}_AutoFitColumnsBaseline.xlsx");
    }

    private static WorkSheet CreateSheetWithBaseline(string sheetName, double calibration, double baseline, uint startRow)
    {
        var sheet = new WorkSheet(sheetName);

        sheet.AddCell(new(0, startRow), $"Example: {sheetName}", cell => cell
            .WithColor("4472C4")
            .WithFont(font => font.Bold().WithSize(14).WithColor("FFFFFF")));
        sheet.MergeCells(0, startRow, 2, startRow);

        sheet.AddCell(new(0, startRow + 1), $"Calibration: {calibration:F1}, Baseline: {baseline:+0.0;-0.0;0.0}", cell => cell
            .WithFont(font => font.Italic())
            .WithColor("D9E1F2"));
        sheet.MergeCells(0, startRow + 1, 2, startRow + 1);

        for (uint col = 0; col < Headers.Length; col++)
            sheet.AddCell(new(col, startRow + 3), Headers[col], cell => cell
                .WithFont(font => font.Bold())
                .WithColor("E7E6E6"));

        for (uint row = 0; row < Data.Length; row++)
            for (uint col = 0; col < Data[row].Length; col++)
            {
                var value = Data[row][col] switch
                {
                    string s => new CellValue(s),
                    decimal d => new CellValue(d),
                    int i => new CellValue(i),
                    _ => new CellValue(Data[row][col].ToString() ?? string.Empty)
                };

                sheet.AddCell(new(col, startRow + 4 + row), value, null);
            }

        if (baseline == 0.0)
            sheet.AutoFitAllColumns(calibration);
        else
            sheet.AutoFitAllColumns(calibration, baseline);

        return sheet;
    }
}
