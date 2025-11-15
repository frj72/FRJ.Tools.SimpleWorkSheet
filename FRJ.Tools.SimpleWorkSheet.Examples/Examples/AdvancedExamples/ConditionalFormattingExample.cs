using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ConditionalFormattingExample : IExample
{
    public string Name => "Conditional Formatting";
    public string Description => "Dynamic styling based on cell values";

    public void Run()
    {
        var sheet = new WorkSheet("ConditionalFormatting");
        
        sheet.AddCell(0, 0, "Score");
        sheet.AddCell(1, 0, "Status");
        
        var scores = new[] { 15, 45, 65, 85, 95 };
        
        for (uint i = 0; i < scores.Length; i++)
        {
            var score = scores[i];
            var row = i + 1;
            
            var color = score switch
            {
                < 30 => "FF0000",
                < 70 => "FFFF00",
                _ => "00FF00"
            };
            
            var status = score switch
            {
                < 30 => "Fail",
                < 70 => "Pass",
                _ => "Excellent"
            };
            
            sheet.AddCell(0, row, score, cell => cell
                .WithColor(color)
                .WithFont(font => font.Bold()));
            
            sheet.AddCell(1, row, status, cell => cell
                .WithColor(color));
        }
        
        ExampleRunner.SaveWorkSheet(sheet, "16_ConditionalFormatting.xlsx");
    }
}