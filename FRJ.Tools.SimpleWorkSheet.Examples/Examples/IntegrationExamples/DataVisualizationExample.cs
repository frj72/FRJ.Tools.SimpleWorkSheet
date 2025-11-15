using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.IntegrationExamples;

public class DataVisualizationExample : IExample
{
    public string Name => "Data Visualization";
    public string Description => "Simple chart-like data presentation";

    private static readonly string[] SourceArray = ["Month", "Sales", "Target", "Achievement"];

    public void Run()
    {
        var sheet = new WorkSheet("DataVisualization");
        
        sheet.AddCell(0, 0, "Sales Report - Q1 2025", cell => cell
            .WithFont(font => font.WithSize(16).Bold())
            .WithColor("2E75B6"));
        
        var headers = SourceArray.Select(h => new CellValue(h));
        
        sheet.AddRow(2, 0, headers, cell => cell
            .WithColor("2E75B6")
            .WithFont(font => font.WithSize(12).WithColor("FFFFFF").Bold()));
        
        var monthData = new[]
        {
            new { Month = "January", Sales = 85000, Target = 80000 },
            new { Month = "February", Sales = 92000, Target = 90000 },
            new { Month = "March", Sales = 78000, Target = 85000 }
        };
        
        for (uint i = 0; i < monthData.Length; i++)
        {
            var row = i + 3;
            var data = monthData[i];
            var achievement = (double)data.Sales / data.Target;
            var achievementPercent = $"{achievement * 100:F1}%";
            
            var bgColor = achievement >= 1.0 ? "D4EDDA" : achievement >= 0.9 ? "FFF3CD" : "F8D7DA";
            var statusColor = achievement >= 1.0 ? "155724" : achievement >= 0.9 ? "856404" : "721C24";
            
            sheet.AddCell(0, row, data.Month);
            sheet.AddCell(1, row, data.Sales, cell => cell.WithFormatCode("$#,##0"));
            sheet.AddCell(2, row, data.Target, cell => cell.WithFormatCode("$#,##0"));
            sheet.AddCell(3, row, achievementPercent, cell => cell
                .WithColor(bgColor)
                .WithFont(font => font
                    .WithColor(statusColor)
                    .Bold()));
        }
        
        var totalSales = monthData.Sum(m => m.Sales);
        var totalTarget = monthData.Sum(m => m.Target);
        var overallAchievement = (double)totalSales / totalTarget;
        
        sheet.AddCell(0, 7, "TOTAL", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(1, 7, totalSales, cell => cell
            .WithFont(font => font.Bold())
            .WithFormatCode("$#,##0"));
        sheet.AddCell(2, 7, totalTarget, cell => cell
            .WithFont(font => font.Bold())
            .WithFormatCode("$#,##0"));
        sheet.AddCell(3, 7, $"{overallAchievement * 100:F1}%", cell => cell
            .WithFont(font => font.Bold())
            .WithColor(overallAchievement >= 1.0 ? "00FF00" : "FFAA00"));
        
        ExampleRunner.SaveWorkSheet(sheet, "20_DataVisualization.xlsx");
    }
}