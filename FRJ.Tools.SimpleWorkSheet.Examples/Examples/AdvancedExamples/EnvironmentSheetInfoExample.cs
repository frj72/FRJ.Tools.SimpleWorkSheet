using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class EnvironmentSheetInfoExample : IExample
{
    public string Name => "Environment Sheet Info - Width Estimation";
    public string Description => "Demonstrates using EnvironmentSheetInfo.GetWidth to estimate text width before adding cells";


    public int ExampleNumber { get; }

    public EnvironmentSheetInfoExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("Width Estimation");

        sheet.AddCell(new(0, 0), "Font", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4").WithFont(f => f.WithColor("FFFFFF"))));
        sheet.AddCell(new(1, 0), "Size", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4").WithFont(f => f.WithColor("FFFFFF"))));
        sheet.AddCell(new(2, 0), "Bold", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4").WithFont(f => f.WithColor("FFFFFF"))));
        sheet.AddCell(new(3, 0), "Italic", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4").WithFont(f => f.WithColor("FFFFFF"))));
        sheet.AddCell(new(4, 0), "Text", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4").WithFont(f => f.WithColor("FFFFFF"))));
        sheet.AddCell(new(5, 0), "Estimated Width", cell => cell
            .WithFont(font => font.Bold())
            .WithStyle(style => style.WithFillColor("4472C4").WithFont(f => f.WithColor("FFFFFF"))));

        var testCases = new[]
        {
            (Font: "Aptos Narrow", Size: 12, Bold: false, Italic: false, Text: "Hello World"),
            (Font: "Aptos Narrow", Size: 14, Bold: false, Italic: false, Text: "Hello World"),
            (Font: "Aptos Narrow", Size: 12, Bold: true, Italic: false, Text: "Hello World"),
            (Font: "Aptos Narrow", Size: 12, Bold: false, Italic: true, Text: "Hello World"),
            (Font: "Aptos Narrow", Size: 12, Bold: true, Italic: true, Text: "Hello World"),
            (Font: "Arial", Size: 12, Bold: false, Italic: false, Text: "Hello World"),
            (Font: "Calibri", Size: 12, Bold: false, Italic: false, Text: "Hello World"),
            (Font: "Aptos Narrow", Size: 12, Bold: false, Italic: false, Text: "Short"),
            (Font: "Aptos Narrow", Size: 12, Bold: false, Italic: false, Text: "This is a much longer text string"),
            (Font: "Aptos Narrow", Size: 20, Bold: true, Italic: false, Text: "Large Bold Text")
        };

        uint row = 1;
        foreach (var testCase in testCases)
        {
            var width = EnvironmentSheetInfo.GetWidth(testCase.Font, testCase.Size, testCase.Text, testCase.Bold, testCase.Italic);

            sheet.AddCell(new(0, row), testCase.Font, null);
            sheet.AddCell(new(1, row), testCase.Size, null);
            sheet.AddCell(new(2, row), testCase.Bold ? "Yes" : "No", null);
            sheet.AddCell(new(3, row), testCase.Italic ? "Yes" : "No", null);
            sheet.AddCell(new(4, row), testCase.Text, cell => cell
                .WithFont(font => 
                {
                    font.WithName(testCase.Font).WithSize(testCase.Size);
                    if (testCase.Bold)
                        font.Bold();
                    if (testCase.Italic)
                        font.Italic();
                }));
            sheet.AddCell(new(5, row), Math.Round(width, 2), null);

            row++;
        }

        sheet.AddCell(new(0, row + 1), "Practical Use Case", cell => cell
            .WithFont(font => font.Bold().WithSize(12))
            .WithStyle(style => style.WithFillColor("DDDDDD")));

        row += 2;

        sheet.AddCell(new(0, row), "Text to Measure:", cell => cell.WithFont(font => font.Bold()));
        var practicalText = "Setting Column Width Based on Content";
        sheet.AddCell(new(1, row), practicalText, null);

        row++;

        var estimatedWidth = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, practicalText);
        sheet.AddCell(new(0, row), "Estimated Width:", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), Math.Round(estimatedWidth, 2), null);

        row++;

        sheet.AddCell(new(0, row), "Applied Width:", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(1, row), "Column B is set to the estimated width", null);
        sheet.SetColumnWidth(1, estimatedWidth);

        sheet.SetColumnWidth(0, 20.0);
        sheet.SetColumnWidth(4, 35.0);
        sheet.SetColumnWidth(5, 18.0);

        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_EnvironmentSheetInfo.xlsx");
    }
}
