using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ChartColorPaletteExample : IExample
{
    public string Name => "Chart Color Palette";
    public string Description => "Shows a simple palette of chart series colors using Colors";


    public int ExampleNumber { get; }

    public ChartColorPaletteExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Color Palette");

        sheet.AddCell(new(0, 0), "Index", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));
        sheet.AddCell(new(1, 0), "Value", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4")));

        var values = new[] { 10, 20, 30, 40, 50, 60, 70, 80 };
        for (uint i = 0; i < values.Length; i++)
        {
            sheet.AddCell(new(0, i + 1), (int)(i + 1), null);
            sheet.AddCell(new(1, i + 1), values[i], null);
        }

        var seriesColors = new[]
        {
            Colors.Red,
            Colors.Green,
            Colors.Blue,
            Colors.VividOrange,
            Colors.MediumPurple,
            Colors.Yellow,
            Colors.Cyan,
            Colors.Magenta
        };

        var startRow = 12U;

        for (var seriesIndex = 0; seriesIndex < seriesColors.Length; seriesIndex++)
        {
            var valuesRange = CellRange.FromBounds(1, 1, 1, (uint)values.Length);

            var chart = LineChart.Create()
                .WithTitle($"Color {seriesIndex + 1}")
                .WithDataRange(valuesRange)
                .WithPosition((uint)(seriesIndex * 4), startRow, (uint)(seriesIndex * 4 + 3), startRow + 10)
                .WithMarkers(LineChartMarkerStyle.Circle)
                .WithSeriesColor(seriesColors[seriesIndex]);

            sheet.AddChart(chart);
        }

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_ChartColorPalette.xlsx");
    }
}
