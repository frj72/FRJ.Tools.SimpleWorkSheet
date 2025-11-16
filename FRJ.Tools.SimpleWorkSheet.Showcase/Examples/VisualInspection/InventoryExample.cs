using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.VisualInspection;

public class InventoryExample : IShowcase
{
    public string Name => "Inventory Tracker";
    public string Description => "Inventory with validation and hyperlinks";
    public string Category => "Visual Inspection";

    public void Run()
    {
        var sheet = new WorkSheet("Inventory");
        
        var headers = new[] { "SKU", "Product", "Quantity", "Price", "Value", "Supplier Website" };
        for (uint col = 0; col < (uint)headers.Length; col++)
            sheet.AddCell(new(col, 0), headers[col], cell => cell.WithFont(f => f.Bold()).WithColor("4472C4"));
        
        var items = new[]
        {
            ("SKU001", "Widget A", 150, 12.50m, "https://supplier-a.com"),
            ("SKU002", "Widget B", 75, 25.00m, "https://supplier-b.com"),
            ("SKU003", "Widget C", 200, 8.75m, "https://supplier-c.com"),
            ("SKU004", "Widget D", 50, 45.00m, "https://supplier-d.com")
        };
        
        uint row = 1;
        foreach (var (sku, product, qty, price, url) in items)
        {
            sheet.AddCell(new(0, row), sku);
            sheet.AddCell(new(1, row), product);
            sheet.AddCell(new(2, row), qty);
            sheet.AddCell(new(3, row), price, cell => cell.WithFormatCode("$#,##0.00"));
            sheet.AddCell(new(4, row), qty * price, cell => cell.WithFormatCode("$#,##0.00"));
            sheet.AddCell(new(5, row), url, cell => cell.WithHyperlink(url, "Visit Supplier"));
            row++;
        }
        
        for (uint col = 0; col < 6; col++)
            sheet.SetColumnWith(col, 15.0);
        
        var workbook = new WorkBook("Inventory", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_16_Inventory.xlsx");
    }
}
