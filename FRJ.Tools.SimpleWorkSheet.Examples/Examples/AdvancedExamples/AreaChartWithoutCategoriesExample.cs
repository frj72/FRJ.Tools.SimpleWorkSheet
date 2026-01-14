using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class AreaChartWithoutCategoriesExample : IExample
{
    public string Name => "Area Chart Without Categories";
    public string Description => "Demonstrates area charts without explicit x-axis categories";

    public void Run()
    {
        var sheet = new WorkSheet("Growth Data");

        var growthData = new[] { 20, 35, 45, 50, 65, 75, 80, 90, 95, 100 };
        
        for (uint i = 0; i < growthData.Length; i++)
            sheet.AddCell(new(0, i), new(growthData[i]), null);

        var valuesRange = CellRange.FromBounds(0, 0, 0, (uint)growthData.Length - 1);

        var areaChart = AreaChart.Create()
            .WithTitle("Growth Over Time")
            .WithDataRange(valuesRange)
            .WithPosition(2, 0, 10, 15)
            .WithValueAxisTitle("Growth Index");

        sheet.AddChart(areaChart);

        ExampleRunner.SaveWorkSheet(sheet, "093_AreaChartWithoutCategories.xlsx");
    }
}
