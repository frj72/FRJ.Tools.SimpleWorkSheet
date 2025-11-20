using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.EdgeCases;

public class ColorPaletteExample : IShowcase
{
    public string Name => "Color Palette (Predefined + Custom)";
    public string Description => "Shows all predefined colors plus custom hex colors";
    public string Category => "Edge Cases";

    public void Run()
    {
        var sheet = new WorkSheet("Colors");
        
        sheet.AddCell(new(0, 0), "Color Name", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 0), "Hex", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(2, 0), "Sample", cell => cell.WithFont(f => f.Bold()));
        
        var colors = new[]
        {
            ("Black", Colors.Black),
            ("White", Colors.White),
            ("Red", Colors.Red),
            ("Green", Colors.Green),
            ("Blue", Colors.Blue),
            ("Yellow", Colors.Yellow),
            ("Cyan", Colors.Cyan),
            ("Magenta", Colors.Magenta),
            ("Orange", "FFA500"),
            ("Purple", "800080"),
            ("Pink", "FFC0CB"),
            ("Brown", "A52A2A"),
            ("Gray", "808080"),
            ("Light Blue", "ADD8E6"),
            ("Dark Green", "006400")
        };
        
        uint row = 1;
        foreach (var (name, hex) in colors)
        {
            sheet.AddCell(new(0, row), name, null);
            sheet.AddCell(new(1, row), hex, null);
            sheet.AddCell(new(2, row), "        ", cell => cell.WithColor(hex));
            row++;
        }
        
        sheet.SetColumnWidth(0, 15.0);
        sheet.SetColumnWidth(1, 10.0);
        sheet.SetColumnWidth(2, 15.0);
        
        var workbook = new WorkBook("Colors", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_07_ColorPalette.xlsx");
    }
}
