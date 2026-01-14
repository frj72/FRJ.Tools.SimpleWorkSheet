using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class AreaChartExample : IExample
{
    public string Name => "Area Chart";
    public string Description => "Area charts showing trends over time";

    private static readonly string[] Months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    private static readonly int[] Sales = [12000, 15000, 13500, 18000, 16500, 21000];

    public void Run()
    {
        var sheet = new WorkSheet("Sales Data");

        sheet.AddCell(new(0, 0), "Month", null);
        sheet.AddCell(new(1, 0), "Sales", null);

        for (uint i = 0; i < Months.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), Months[i], null);
            sheet.AddCell(new(1, i + 1), Sales[i], null);
        }

        var chart = AreaChart.Create()
            .WithTitle("Monthly Sales Trend")
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, (uint)Months.Length),
                CellRange.FromBounds(1, 1, 1, (uint)Months.Length))
            .WithPosition(3, 0, 12, 15)
            .WithSize(5000000, 3000000);

        sheet.AddChart(chart);

        ExampleRunner.SaveWorkSheet(sheet, "051_AreaChart.xlsx");
    }
}
