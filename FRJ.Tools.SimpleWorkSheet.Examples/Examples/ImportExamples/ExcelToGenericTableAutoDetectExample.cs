using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ExcelToGenericTableAutoDetectExample : IExample
{
    public string Name => "ExcelToGenericTableAutoDetect";
    public string Description => "Import mixed workbook using HeaderMode.AutoDetect";

    public int ExampleNumber { get; }

    public ExcelToGenericTableAutoDetectExample(int exampleNumber) => ExampleNumber = exampleNumber;

    public void Run()
    {
        var salesSheet = new WorkSheet("Sales");

        salesSheet.AddCell(new(0, 0), "Product", cell => cell.WithFont(f => f.Bold()));
        salesSheet.AddCell(new(1, 0), "Price", cell => cell.WithFont(f => f.Bold()));
        salesSheet.AddCell(new(2, 0), "Quantity", cell => cell.WithFont(f => f.Bold()));

        salesSheet.AddCell(new(0, 1), "Widget A", null);
        salesSheet.AddCell(new(1, 1), 29.99m, null);
        salesSheet.AddCell(new(2, 1), 150L, null);

        salesSheet.AddCell(new(0, 2), "Widget B", null);
        salesSheet.AddCell(new(1, 2), 49.99m, null);
        salesSheet.AddCell(new(2, 2), 200L, null);

        salesSheet.AddCell(new(0, 3), "Widget C", null);
        salesSheet.AddCell(new(1, 3), 19.99m, null);
        salesSheet.AddCell(new(2, 3), 300L, null);

        salesSheet.AddCell(new(0, 4), "Widget D", null);
        salesSheet.AddCell(new(1, 4), 39.99m, null);
        salesSheet.AddCell(new(2, 4), 175L, null);

        salesSheet.AddCell(new(0, 5), "Widget E", null);
        salesSheet.AddCell(new(1, 5), 59.99m, null);
        salesSheet.AddCell(new(2, 5), 125L, null);

        var matrixSheet = new WorkSheet("Matrix");

        matrixSheet.AddCell(new(0, 0), 10.5m, null);
        matrixSheet.AddCell(new(1, 0), 12.3m, null);
        matrixSheet.AddCell(new(2, 0), 15.7m, null);
        matrixSheet.AddCell(new(3, 0), 18.2m, null);

        matrixSheet.AddCell(new(0, 1), 11.8m, null);
        matrixSheet.AddCell(new(1, 1), 13.5m, null);
        matrixSheet.AddCell(new(2, 1), 16.9m, null);
        matrixSheet.AddCell(new(3, 1), 19.4m, null);

        matrixSheet.AddCell(new(0, 2), 9.7m, null);
        matrixSheet.AddCell(new(1, 2), 11.2m, null);
        matrixSheet.AddCell(new(2, 2), 14.6m, null);
        matrixSheet.AddCell(new(3, 2), 17.1m, null);

        matrixSheet.AddCell(new(0, 3), 12.4m, null);
        matrixSheet.AddCell(new(1, 3), 14.1m, null);
        matrixSheet.AddCell(new(2, 3), 17.5m, null);
        matrixSheet.AddCell(new(3, 3), 20.0m, null);

        var workbook = new WorkBook("MixedData", [salesSheet, matrixSheet]);

        var tempPath = Path.Combine(Path.GetTempPath(), "mixed_data_temp.xlsx");
        workbook.SaveToFile(tempPath);

        var tables = WorkBookReader.LoadAsGenericTables(tempPath, HeaderMode.AutoDetect);

        var salesOutput = WorksheetBuilder.FromGenericTable(tables["Sales"])
            .WithSheetName("Sales_Imported")
            .WithHeaderStyle(style => style.WithFont(font => font.Bold()))
            .AutoFitAllColumns()
            .Build();

        var matrixOutput = WorksheetBuilder.FromGenericTable(tables["Matrix"])
            .WithSheetName("Matrix_Imported")
            .WithHeaderStyle(style => style.WithFont(font => font.Bold()))
            .AutoFitAllColumns()
            .Build();

        var finalWorkbook = new WorkBook("ImportedData", [salesOutput, matrixOutput]);
        
        ExampleRunner.SaveWorkBook(finalWorkbook, $"{ExampleNumber:000}_{Name}.xlsx");

        File.Delete(tempPath);
    }
}
