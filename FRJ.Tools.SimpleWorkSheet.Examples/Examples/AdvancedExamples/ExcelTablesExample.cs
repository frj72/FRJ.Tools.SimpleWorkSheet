using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ExcelTablesExample : IExample
{
    public string Name => "Excel Tables";
    public string Description => "Native Excel tables with automatic filters";

    private static readonly string[] Headers = ["Product", "Category", "Price", "Quantity", "Total"];
    private static readonly object[][] Data =
    [
        ["Laptop", "Electronics", 1299.99m, 15, 19499.85m],
        ["Mouse", "Electronics", 24.99m, 120, 2998.80m],
        ["Desk", "Furniture", 459.99m, 8, 3679.92m],
        ["Chair", "Furniture", 189.99m, 25, 4749.75m],
        ["Monitor", "Electronics", 329.99m, 18, 5939.82m],
        ["Lamp", "Furniture", 45.99m, 50, 2299.50m],
        ["Keyboard", "Electronics", 79.99m, 40, 3199.60m],
        ["Bookshelf", "Furniture", 199.99m, 12, 2399.88m],
        ["Webcam", "Electronics", 89.99m, 30, 2699.70m],
        ["Standing Desk", "Furniture", 699.99m, 5, 3499.95m]
    ];

    public void Run()
    {
        var sheet = new WorkSheet("Inventory");

        sheet.AddCell(new(0, 0), "Product Inventory", cell => cell
            .WithFont(font => font.Bold().WithSize(14))
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.MergeCells(0, 0, 4, 0);

        for (uint col = 0; col < Headers.Length; col++)
            sheet.AddCell(new(col, 1), Headers[col], cell => cell
                .WithFont(font => font.Bold())
                .WithStyle(style => style.WithFillColor("D9E1F2")));

        for (uint row = 0; row < Data.Length; row++)
        for (uint col = 0; col < Data[row].Length; col++)
        {
            var value = Data[row][col] switch
            {
                string s => new(s),
                decimal d => new(d),
                int i => new(i),
                _ => new Components.SimpleCell.CellValue(Data[row][col].ToString() ?? string.Empty)
            };
            sheet.AddCell(new(col, row + 2), value);
        }

        sheet.AddTable("InventoryTable", 0, 1, 4, (uint)(Data.Length + 1));

        sheet.AutoFitAllColumns();

        ExampleRunner.SaveWorkSheet(sheet, "49_ExcelTables.xlsx");
    }
}
