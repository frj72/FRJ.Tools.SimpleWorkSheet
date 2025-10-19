using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples;

public class FontVariationsExample : IExample
{
    public string Name => "Font Variations";
    public string Description => "Bold, italic, size, color, name changes";

    public void Run()
    {
        var sheet = new WorkSheet("FontVariations");
        
        sheet.AddCell(0, 0, "Bold Text", cell => cell
            .WithFont(font => font.Bold()));
        
        sheet.AddCell(0, 1, "Italic Text", cell => cell
            .WithFont(font => font.Italic()));
        
        sheet.AddCell(0, 2, "Large Text", cell => cell
            .WithFont(font => font.WithSize(20)));
        
        sheet.AddCell(0, 3, "Red Text", cell => cell
            .WithFont(font => font.WithColor("FF0000")));
        
        sheet.AddCell(0, 4, "Calibri Font", cell => cell
            .WithFont(font => font.WithName("Calibri")));
        
        sheet.AddCell(0, 5, "Bold Italic Underline", cell => cell
            .WithFont(font => font
                .Bold()
                .Italic()
                .Underline()));
        
        ExampleRunner.SaveWorkSheet(sheet, "06_FontVariations.xlsx");
    }
}

public class BackgroundColorsExample : IExample
{
    public string Name => "Background Colors";
    public string Description => "Different fill colors";

    public void Run()
    {
        var sheet = new WorkSheet("BackgroundColors");
        
        sheet.AddCell(0, 0, "Red Background", cell => cell.WithColor("FF0000"));
        sheet.AddCell(0, 1, "Green Background", cell => cell.WithColor("00FF00"));
        sheet.AddCell(0, 2, "Blue Background", cell => cell.WithColor("0000FF"));
        sheet.AddCell(0, 3, "Yellow Background", cell => cell.WithColor("FFFF00"));
        sheet.AddCell(0, 4, "Gray Background", cell => cell.WithColor("CCCCCC"));
        sheet.AddCell(0, 5, "Light Blue Background", cell => cell.WithColor("ADD8E6"));
        
        ExampleRunner.SaveWorkSheet(sheet, "07_BackgroundColors.xlsx");
    }
}

public class BorderStylesExample : IExample
{
    public string Name => "Border Styles";
    public string Description => "Various border configurations";

    public void Run()
    {
        var sheet = new WorkSheet("BorderStyles");
        
        var thinBorders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        sheet.AddCell(0, 0, "All Borders", cell => cell.WithBorders(thinBorders));
        
        var topBottomBorders = CellBorders.Create(
            null,
            null,
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        sheet.AddCell(0, 1, "Top/Bottom Only", cell => cell.WithBorders(topBottomBorders));
        
        var leftBorder = CellBorders.Create(
            CellBorder.Create(Colors.Red, CellBorderStyle.Thin),
            null,
            null,
            null);
        
        sheet.AddCell(0, 2, "Left Border Red", cell => cell.WithBorders(leftBorder));
        
        ExampleRunner.SaveWorkSheet(sheet, "08_BorderStyles.xlsx");
    }
}

public class CompleteStylingExample : IExample
{
    public string Name => "Complete Styling";
    public string Description => "Combining all style elements";

    public void Run()
    {
        var sheet = new WorkSheet("CompleteStyling");
        
        var borders = CellBorders.Create(
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin),
            CellBorder.Create(Colors.Black, CellBorderStyle.Thin));
        
        sheet.AddCell(0, 0, 123.456m, cell => cell
            .WithStyle(style => style
                .WithFillColor("E0E0E0")
                .WithFont(font => font
                    .WithSize(14)
                    .WithName("Calibri")
                    .WithColor("0000FF")
                    .Bold()
                    .Italic())
                .WithBorders(borders)
                .WithFormatCode("0.00")));
        
        sheet.AddCell(0, 1, "Header Style", cell => cell
            .WithColor("4472C4")
            .WithFont(font => font
                .WithSize(16)
                .WithColor("FFFFFF")
                .Bold())
            .WithBorders(borders));
        
        ExampleRunner.SaveWorkSheet(sheet, "09_CompleteStyling.xlsx");
    }
}
