using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class GenericTableAdvancedExample : IExample
{
    public string Name => "075_GenericTableAdvanced";
    public string Description => "Advanced Generic Table with filtering, parsers, and conditional styles";

    public void Run()
    {
        var table = GenericTable.Create("Product", "Category", "Price", "Stock", "LastUpdate");
        
        table.AddRow(
            new CellValue("Laptop"), 
            new CellValue("Electronics"), 
            new CellValue(1200m),
            new CellValue(15), 
            new CellValue(new DateTime(2024, 11, 1)));
            
        table.AddRow(
            new CellValue("Mouse"), 
            new CellValue("Electronics"), 
            new CellValue(25m),
            new CellValue(150), 
            new CellValue(new DateTime(2024, 11, 15)));
            
        table.AddRow(
            new CellValue("Desk"), 
            new CellValue("Furniture"), 
            new CellValue(450m),
            new CellValue(8), 
            new CellValue(new DateTime(2024, 10, 20)));
            
        table.AddRow(
            new CellValue("Chair"), 
            new CellValue("Furniture"), 
            new CellValue(180m),
            new CellValue(25), 
            new CellValue(new DateTime(2024, 11, 10)));
        
        var sheet = WorksheetBuilder.FromGenericTable(table)
            .WithSheetName("Inventory")
            .WithHeaderStyle(style => style
                .WithFont(font => font.Bold().WithSize(12).WithColor(Colors.White))
                .WithFillColor("FF4472C4"))
            .WithColumnOrder("Product", "Category", "Stock", "Price", "LastUpdate")
            .WithColumnParser("Price", value => new(value.Value.AsT0 * 1.1m))
            .WithNumberFormat("Price", NumberFormat.Float2)
            .WithDateFormat(DateFormat.IsoDate)
            .WithConditionalStyle("Stock",
                value => value.Value.AsT1 < 10,
                style => style.WithFillColor("FFFFCCCC"))
            .WithConditionalStyle("Price",
                value => value.Value.AsT0 > 500,
                style => style.WithFillColor("FFC6EFCE"))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{Name}.xlsx");
    }
}
