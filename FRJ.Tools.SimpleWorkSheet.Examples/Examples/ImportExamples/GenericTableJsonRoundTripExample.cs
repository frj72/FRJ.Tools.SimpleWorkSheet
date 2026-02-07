using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class GenericTableJsonRoundTripExample : IExample
{
    public string Name => "GenericTableJsonRoundTrip";
    public string Description => "Demonstrates GenericTable export to JSON and round-trip import";
    public int ExampleNumber { get; }

    public GenericTableJsonRoundTripExample(int exampleNumber) => ExampleNumber = exampleNumber;

    public void Run()
    {
        var table = GenericTable.Create("Name", "Salary", "EmployeeID", "HireDate", "Department");
        table.AddRow(
            new CellValue("Alice"),
            new CellValue(75000.50m),
            new CellValue(1001L),
            new CellValue(new DateTime(2020, 1, 15)),
            new CellValue("Engineering")
        );
        table.AddRow(
            new CellValue("Bob"),
            new CellValue(68000m),
            new CellValue(1002L),
            new CellValue(new DateTime(2019, 6, 20)),
            new CellValue("Sales")
        );
        table.AddRow(
            new CellValue("Charlie"),
            null,
            new CellValue(1003L),
            new CellValue(new DateTime(2021, 3, 10)),
            new CellValue("Management")
        );
        table.AddRow(
            new CellValue("Diana"),
            new CellValue(80000m),
            new CellValue(1004L),
            new CellValue(new DateTime(2022, 8, 5)),
            null
        );

        var originalSheet = WorksheetBuilder.FromGenericTable(table)
            .WithSheetName("Original")
            .WithHeaderStyle(style => style
                .WithFont(font => font.Bold())
                .WithFillColor("FFD3D3D3"))
            .AutoFitAllColumns()
            .Build();

        var jsonContent = GenericTableToJsonConverter.ToJson(table);
        using var jsonDoc = JsonDocument.Parse(jsonContent);
        var importedTable = JsonToGenericTableConverter.Convert(jsonDoc.RootElement);

        var roundTripSheet = WorksheetBuilder.FromGenericTable(importedTable)
            .WithSheetName("FromJsonRoundTrip")
            .WithHeaderStyle(style => style
                .WithFont(font => font.Bold())
                .WithFillColor("FFD3D3D3"))
            .AutoFitAllColumns()
            .Build();

        var workbook = new WorkBook("EmployeeData", [originalSheet, roundTripSheet]);

        ExampleRunner.SaveWorkBook(workbook, $"{ExampleNumber:000}_{Name}.xlsx");
    }
}
