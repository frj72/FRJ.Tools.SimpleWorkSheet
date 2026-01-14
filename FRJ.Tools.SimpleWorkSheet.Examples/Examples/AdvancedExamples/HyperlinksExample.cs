using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class HyperlinksExample : IExample
{
    public string Name => "Hyperlinks";
    public string Description => "Demonstrates external URLs and email links";

    public void Run()
    {
        var sheet = new WorkSheet("Hyperlinks");

        sheet.AddCell(new(0, 0), "Type", null);
        sheet.AddCell(new(1, 0), "Link", null);

        sheet.AddCell(new(0, 1), "Website", null);
        sheet.AddCell(new(1, 1), "Visit Example.com", cell => cell
            .WithHyperlink("https://example.com", "Click to visit our website")
            .WithFont(font => font.WithColor("0000FF")));

        sheet.AddCell(new(0, 2), "GitHub", null);
        sheet.AddCell(new(1, 2), "View Repository", cell => cell
            .WithHyperlink("https://github.com", null)
            .WithFont(font => font.WithColor("0000FF")));

        sheet.AddCell(new(0, 3), "Email", null);
        sheet.AddCell(new(1, 3), "Send Email", cell => cell
            .WithHyperlink("mailto:contact@example.com", "Send us an email")
            .WithFont(font => font.WithColor("0000FF")));

        sheet.AddCell(new(0, 4), "Documentation", null);
        sheet.AddCell(new(1, 4), "Read Docs", cell => cell
            .WithHyperlink("https://docs.example.com/guide", null)
            .WithFont(font => font.WithColor("0000FF").Bold()));

        sheet.AddCell(new(0, 6), "Styled Link", cell => cell
            .WithFont(font => font.Bold()));
        sheet.AddCell(new(1, 6), "Centered Hyperlink", cell => cell
            .WithHyperlink("https://example.com", null)
            .WithFont(font => font.WithColor("0000FF").Italic())
            .WithStyle(style => style
                .WithHorizontalAlignment(HorizontalAlignment.Center)
                .WithFillColor("F2F2F2")));

        ExampleRunner.SaveWorkSheet(sheet, "034_Hyperlinks.xlsx");
    }
}
