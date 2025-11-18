using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class StackedAreaChartExample : IExample
{
    public string Name => "Stacked Area Chart";
    public string Description => "Stacked area charts showing cumulative values";

    private static readonly string[] Months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    private static readonly int[] ProductA = [5000, 6000, 5500, 7000, 6500, 8000];
    private static readonly int[] ProductB = [3000, 4000, 3500, 5000, 4500, 6000];
    private static readonly int[] ProductC = [2000, 2500, 2000, 3000, 2800, 3500];

    public void Run()
    {
        var sheet = new WorkSheet("Sales Data");

        sheet.AddCell(new(0, 0), "Month");
        sheet.AddCell(new(1, 0), "Product A");
        sheet.AddCell(new(2, 0), "Product B");
        sheet.AddCell(new(3, 0), "Product C");

        for (uint i = 0; i < Months.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), Months[i]);
            sheet.AddCell(new(1, i + 1), ProductA[i]);
            sheet.AddCell(new(2, i + 1), ProductB[i]);
            sheet.AddCell(new(3, i + 1), ProductC[i]);
        }

        var chart = AreaChart.Create()
            .WithTitle("Product Sales - Stacked")
            .WithPosition(5, 0, 14, 15)
            .WithSize(5000000, 3000000)
            .WithStacked(true)
            .AddSeries("Product A", CellRange.FromBounds(1, 1, 1, (uint)Months.Length))
            .AddSeries("Product B", CellRange.FromBounds(2, 1, 2, (uint)Months.Length))
            .AddSeries("Product C", CellRange.FromBounds(3, 1, 3, (uint)Months.Length));

        sheet.AddChart(chart);

        ExampleRunner.SaveWorkSheet(sheet, "52_StackedAreaChart.xlsx");
    }
}
