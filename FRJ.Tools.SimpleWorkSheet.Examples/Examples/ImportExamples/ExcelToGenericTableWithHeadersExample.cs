using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ExcelToGenericTableWithHeadersExample : IExample
{
    public string Name => "ExcelToGenericTableWithHeaders";
    public string Description => "Import employee data with headers using HeaderMode.FirstRow";

    public int ExampleNumber { get; }

    public ExcelToGenericTableWithHeadersExample(int exampleNumber) => ExampleNumber = exampleNumber;

    public void Run()
    {
        var inputSheet = new WorkSheet("Employees");

        inputSheet.AddCell(new(0, 0), "Name", cell => cell.WithFont(f => f.Bold()));
        inputSheet.AddCell(new(1, 0), "Department", cell => cell.WithFont(f => f.Bold()));
        inputSheet.AddCell(new(2, 0), "Salary", cell => cell.WithFont(f => f.Bold()));

        inputSheet.AddCell(new(0, 1), "Alice Johnson", null);
        inputSheet.AddCell(new(1, 1), "Engineering", null);
        inputSheet.AddCell(new(2, 1), 85000m, null);

        inputSheet.AddCell(new(0, 2), "Bob Smith", null);
        inputSheet.AddCell(new(1, 2), "Sales", null);
        inputSheet.AddCell(new(2, 2), 72000m, null);

        inputSheet.AddCell(new(0, 3), "Carol White", null);
        inputSheet.AddCell(new(1, 3), "Engineering", null);
        inputSheet.AddCell(new(2, 3), 91000m, null);

        inputSheet.AddCell(new(0, 4), "David Brown", null);
        inputSheet.AddCell(new(1, 4), "Management", null);
        inputSheet.AddCell(new(2, 4), 105000m, null);

        inputSheet.AddCell(new(0, 5), "Eve Davis", null);
        inputSheet.AddCell(new(1, 5), "Sales", null);
        inputSheet.AddCell(new(2, 5), 68000m, null);

        inputSheet.AddCell(new(0, 6), "Frank Miller", null);
        inputSheet.AddCell(new(1, 6), "Engineering", null);
        inputSheet.AddCell(new(2, 6), 88000m, null);

        inputSheet.AddCell(new(0, 7), "Grace Wilson", null);
        inputSheet.AddCell(new(1, 7), "Management", null);
        inputSheet.AddCell(new(2, 7), 98000m, null);

        inputSheet.AddCell(new(0, 8), "Henry Moore", null);
        inputSheet.AddCell(new(1, 8), "Sales", null);
        inputSheet.AddCell(new(2, 8), 75000m, null);

        inputSheet.AddCell(new(0, 9), "Iris Taylor", null);
        inputSheet.AddCell(new(1, 9), "Engineering", null);
        inputSheet.AddCell(new(2, 9), 92000m, null);

        inputSheet.AddCell(new(0, 10), "Jack Anderson", null);
        inputSheet.AddCell(new(1, 10), "Sales", null);
        inputSheet.AddCell(new(2, 10), 71000m, null);

        inputSheet.AddCell(new(0, 11), "Karen Thomas", null);
        inputSheet.AddCell(new(1, 11), "Management", null);
        inputSheet.AddCell(new(2, 11), 110000m, null);

        inputSheet.AddCell(new(0, 12), "Leo Jackson", null);
        inputSheet.AddCell(new(1, 12), "Engineering", null);
        inputSheet.AddCell(new(2, 12), 86000m, null);

        var tempPath = Path.Combine(Path.GetTempPath(), "employee_data_temp.xlsx");
        ExampleRunner.SaveWorkSheet(inputSheet, tempPath);

        var tables = WorkBookReader.LoadAsGenericTables(tempPath, HeaderMode.FirstRow);

        var outputSheet = new WorkSheet("ImportedData");

        var table = tables["Employees"];

        for (var col = 0; col < table.ColumnCount; col++)
            outputSheet.AddCell(new((uint)col, 0), table.Headers[col], cell => cell.WithFont(f => f.Bold()));

        for (var row = 0; row < table.RowCount; row++)
        for (var col = 0; col < table.ColumnCount; col++)
        {
            var value = table.GetValue(col, row);
            if (value == null)
                continue;

            if (value.IsString())
                outputSheet.AddCell(new((uint)col, (uint)(row + 1)), value.AsString(), null);
            else if (value.IsDecimal())
                outputSheet.AddCell(new((uint)col, (uint)(row + 1)), value.AsDecimal(), null);
        }

        for (uint col = 0; col < (uint)table.ColumnCount; col++)
            outputSheet.AutoFitColumn(col);

        ExampleRunner.SaveWorkSheet(outputSheet, $"{ExampleNumber:000}_{Name}.xlsx");

        File.Delete(tempPath);
    }
}
