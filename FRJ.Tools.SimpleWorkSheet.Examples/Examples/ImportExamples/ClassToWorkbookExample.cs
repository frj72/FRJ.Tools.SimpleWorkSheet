using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ClassToWorkbookExample : IExample
{
    public string Name => "Class to Workbook";
    public string Description => "Creates a workbook from a C# class instance using FromClass<T>";

    public void Run()
    {
        var products = new[]
        {
            new { Id = 1, Name = "Laptop", Price = 999.99m, InStock = true, Category = "Electronics" },
            new { Id = 2, Name = "Mouse", Price = 24.99m, InStock = true, Category = "Electronics" },
            new { Id = 3, Name = "Keyboard", Price = 79.99m, InStock = false, Category = "Electronics" },
            new { Id = 4, Name = "Monitor", Price = 299.99m, InStock = true, Category = "Electronics" },
            new { Id = 5, Name = "Desk Chair", Price = 199.99m, InStock = true, Category = "Furniture" }
        };

        var workbook = WorkbookBuilder.FromClass(products)
            .WithWorkbookName("Product Catalog")
            .WithDataSheetName("Products")
            .WithHeaderStyle(style => style
                .WithFillColor("4472C4")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithColumnParser("Price", value => new(value.Value.AsT0))
            .WithNumberFormat("Price", NumberFormat.Float2)
            .WithConditionalStyle("InStock",
                value => value.Value.AsT2 == "TRUE",
                style => style.WithFillColor(Colors.Green))
            .WithConditionalStyle("InStock",
                value => value.Value.AsT2 == "FALSE",
                style => style.WithFillColor(Colors.Red))
            .WithFreezeHeaderRow()
            .AutoFitAllColumns()
            .WithChart(chart => chart
                .OnSheet("Price Chart")
                .UseColumns("Name", "Price")
                .AsBarChart()
                .WithTitle("Product Prices")
                .WithCategoryAxisTitle("Product")
                .WithValueAxisTitle("Price ($)"))
            .Build();

        ExampleRunner.SaveWorkBook(workbook, "079_ClassToWorkbook.xlsx");
    }
}
