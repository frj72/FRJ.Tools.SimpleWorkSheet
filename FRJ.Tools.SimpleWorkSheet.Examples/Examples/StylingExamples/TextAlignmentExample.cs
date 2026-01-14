using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;

public class TextAlignmentExample : IExample
{
    public string Name => "Text Alignment";
    public string Description => "Demonstrates horizontal/vertical alignment, rotation, and wrapping";

    public void Run()
    {
        var sheet = new WorkSheet("Alignment");

        sheet.AddCell(new(0, 0), "Horizontal Alignment", null);
        sheet.AddCell(new(0, 1), "Left (default)", cell => cell
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Left)));
        sheet.AddCell(new(0, 2), "Center", cell => cell
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Center)));
        sheet.AddCell(new(0, 3), "Right", cell => cell
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Right)));

        sheet.AddCell(new(2, 0), "Vertical Alignment", null);
        sheet.SetRowHeight(1, 40.0);
        sheet.SetRowHeight(2, 40.0);
        sheet.SetRowHeight(3, 40.0);
        sheet.AddCell(new(2, 1), "Top", cell => cell
            .WithStyle(style => style.WithVerticalAlignment(VerticalAlignment.Top)));
        sheet.AddCell(new(2, 2), "Middle", cell => cell
            .WithStyle(style => style.WithVerticalAlignment(VerticalAlignment.Middle)));
        sheet.AddCell(new(2, 3), "Bottom", cell => cell
            .WithStyle(style => style.WithVerticalAlignment(VerticalAlignment.Bottom)));

        sheet.AddCell(new(4, 0), "Text Rotation", null);
        sheet.AddCell(new(4, 1), "45 degrees", cell => cell
            .WithStyle(style => style.WithTextRotation(45)));
        sheet.AddCell(new(4, 2), "-30 degrees", cell => cell
            .WithStyle(style => style.WithTextRotation(-30)));

        sheet.AddCell(new(6, 0), "Wrap Text", null);
        sheet.SetColumnWidth(6, 15.0);
        sheet.AddCell(new(6, 1), "This is a long text that will wrap to multiple lines", cell => cell
            .WithStyle(style => style.WithWrapText(true)));

        sheet.AddCell(new(0, 5), "Combined", cell => cell
            .WithStyle(style => style
                .WithHorizontalAlignment(HorizontalAlignment.Center)
                .WithVerticalAlignment(VerticalAlignment.Middle)
                .WithFillColor("E7E6E6"))
            .WithFont(font => font.Bold()));
        sheet.SetRowHeight(5, 40.0);

        ExampleRunner.SaveWorkSheet(sheet, "033_TextAlignment.xlsx");
    }
}
