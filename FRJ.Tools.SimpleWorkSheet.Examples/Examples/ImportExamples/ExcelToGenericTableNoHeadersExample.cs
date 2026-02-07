using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class ExcelToGenericTableNoHeadersExample : IExample
{
    public string Name => "ExcelToGenericTableNoHeaders";
    public string Description => "Import numeric matrix without headers using HeaderMode.None";

    public int ExampleNumber { get; }

    public ExcelToGenericTableNoHeadersExample(int exampleNumber) => ExampleNumber = exampleNumber;

    public void Run()
    {
        var inputSheet = new WorkSheet("SensorData");

        inputSheet.AddCell(new(0, 0), 23.5m, null);
        inputSheet.AddCell(new(1, 0), 24.1m, null);
        inputSheet.AddCell(new(2, 0), 22.8m, null);
        inputSheet.AddCell(new(3, 0), 25.3m, null);

        inputSheet.AddCell(new(0, 1), 22.9m, null);
        inputSheet.AddCell(new(1, 1), 23.7m, null);
        inputSheet.AddCell(new(2, 1), 23.2m, null);
        inputSheet.AddCell(new(3, 1), 24.8m, null);

        inputSheet.AddCell(new(0, 2), 24.2m, null);
        inputSheet.AddCell(new(1, 2), 25.0m, null);
        inputSheet.AddCell(new(2, 2), 23.5m, null);
        inputSheet.AddCell(new(3, 2), 24.6m, null);

        inputSheet.AddCell(new(0, 3), 23.1m, null);
        inputSheet.AddCell(new(1, 3), 22.5m, null);
        inputSheet.AddCell(new(2, 3), 24.0m, null);
        inputSheet.AddCell(new(3, 3), 23.8m, null);

        inputSheet.AddCell(new(0, 4), 25.1m, null);
        inputSheet.AddCell(new(1, 4), 24.9m, null);
        inputSheet.AddCell(new(2, 4), 25.5m, null);
        inputSheet.AddCell(new(3, 4), 26.0m, null);

        inputSheet.AddCell(new(0, 5), 24.5m, null);
        inputSheet.AddCell(new(1, 5), 23.9m, null);
        inputSheet.AddCell(new(2, 5), 24.3m, null);
        inputSheet.AddCell(new(3, 5), 25.2m, null);

        var tempPath = Path.Combine(Path.GetTempPath(), "sensor_data_temp.xlsx");
        ExampleRunner.SaveWorkSheet(inputSheet, tempPath);

        var tables = WorkBookReader.LoadAsGenericTables(tempPath);

        var outputSheet = new WorkSheet("ImportedData");

        var table = tables["SensorData"];

        for (var col = 0; col < table.ColumnCount; col++)
            outputSheet.AddCell(new((uint)col, 0), table.Headers[col], cell => cell.WithFont(f => f.Bold()));

        for (var row = 0; row < table.RowCount; row++)
        for (var col = 0; col < table.ColumnCount; col++)
        {
            var value = table.GetValue(col, row);
            if (value != null)
                outputSheet.AddCell(new((uint)col, (uint)(row + 1)), value.AsDecimal(), null);
        }

        for (uint col = 0; col < (uint)table.ColumnCount; col++)
            outputSheet.AutoFitColumn(col);

        ExampleRunner.SaveWorkSheet(outputSheet, $"{ExampleNumber:000}_{Name}.xlsx");

        File.Delete(tempPath);
    }
}
