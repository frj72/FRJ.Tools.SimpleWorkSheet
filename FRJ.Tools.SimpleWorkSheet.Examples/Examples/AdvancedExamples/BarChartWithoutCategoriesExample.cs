using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class BarChartWithoutCategoriesExample : IExample
{
    public string Name => "Bar Chart Without Categories";
    public string Description => "Demonstrates bar charts without explicit x-axis categories";


    public int ExampleNumber { get; }

    public BarChartWithoutCategoriesExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Revenue Data");

        var revenueData = new[] { 85000, 92000, 88000, 95000, 91000, 98000 };
        
        for (uint i = 0; i < revenueData.Length; i++)
            sheet.AddCell(new(0, i), new(revenueData[i]), null);

        var valuesRange = CellRange.FromBounds(0, 0, 0, (uint)revenueData.Length - 1);

        var barChart = BarChart.Create()
            .WithTitle("Revenue Progression")
            .WithDataRange(valuesRange)
            .WithPosition(2, 0, 10, 15)
            .WithOrientation(BarChartOrientation.Vertical)
            .WithValueAxisTitle("Revenue ($)");

        sheet.AddChart(barChart);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_BarChartWithoutCategories.xlsx");
    }
}
