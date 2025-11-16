using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class AutoFitColumnsExample : IExample
{
    public string Name => "Auto-Fit Columns";
    public string Description => "Automatically adjust column widths to fit content";

    private static readonly string[] SourceArrayHeaders = ["Product", "Description", "Price", "Quantity", "Total"];
    private static readonly object[][] SourceArrayData =
    [
        ["Laptop", "High-performance laptop with 16GB RAM", 1299.99m, 5, 6499.95m],
        ["Mouse", "Wireless", 24.99m, 50, 1249.50m],
        ["Monitor", "27-inch 4K Ultra HD Display", 459.99m, 10, 4599.90m],
        ["Keyboard", "Mechanical RGB Gaming Keyboard", 129.99m, 25, 3249.75m],
        ["USB-C Hub", "Multi-port adapter", 39.99m, 100, 3999.00m]
    ];

    public void Run()
    {
        var sheet = new WorkSheet("AutoFitColumns");

        sheet.AddCell(new(0, 0), "Product Inventory Report", cell => cell
            .WithFont(font => font.Bold().WithSize(14))
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.MergeCells(0, 0, 4, 0);

        var headers = SourceArrayHeaders.Select(h => new CellValue(h));
        sheet.AddRow(0, 1, headers, cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("D9E1F2")));

        for (uint row = 0; row < SourceArrayData.Length; row++)
            for (uint col = 0; col < SourceArrayData[row].Length; col++)
            {
                var value = SourceArrayData[row][col] switch
                {
                    string s => new(s),
                    decimal d => new(d),
                    int i => new(i),
                    _ => new CellValue(SourceArrayData[row][col].ToString() ?? string.Empty)
                };

                sheet.AddCell(new(col, row + 2), value);
            }

        sheet.AutoFitAllColumns();

        ExampleRunner.SaveWorkSheet(sheet, "46_AutoFitColumns.xlsx");
    }
}
