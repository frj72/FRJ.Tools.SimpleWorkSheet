using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class LineChartWithoutCategoriesExample : IExample
{
    public string Name => "Line Chart Without Categories";
    public string Description => "Demonstrates line charts without explicit x-axis categories";


    public int ExampleNumber { get; }

    public LineChartWithoutCategoriesExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Sales Data");

        var salesData = new[] { 100000, 120000, 115000, 130000, 125000, 140000 };
        
        for (uint i = 0; i < salesData.Length; i++)
            sheet.AddCell(new(0, i), new(salesData[i]), null);

        var valuesRange = CellRange.FromBounds(0, 0, 0, (uint)salesData.Length - 1);

        var lineChart = LineChart.Create()
            .WithTitle("Sales Trend")
            .WithDataRange(valuesRange)
            .WithPosition(2, 0, 10, 15)
            .WithMarkers(LineChartMarkerStyle.Circle)
            .WithValueAxisTitle("Sales ($)");

        sheet.AddChart(lineChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_LineChartWithoutCategories.xlsx");
    }
}
