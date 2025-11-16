using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.EdgeCases;

public class BorderStylesMatrixExample : IShowcase
{
    public string Name => "Border Styles Matrix (All 14 styles)";
    public string Description => "Grid showing all 14 border styles in action";
    public string Category => "Edge Cases";

    public void Run()
    {
        var sheet = new WorkSheet("BorderStyles");
        
        sheet.AddCell(new(0, 0), "Border Style", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 0), "Sample", cell => cell.WithFont(f => f.Bold()));
        
        var styles = Enum.GetValues<CellBorderStyle>();
        uint row = 1;
        
        foreach (var style in styles)
        {
            sheet.AddCell(new(0, row), style.ToString());
            sheet.AddCell(new(1, row), "Sample Text", cell => cell.WithBorders(
                CellBorders.Create(
                    CellBorder.Create("000000", style),
                    CellBorder.Create("000000", style),
                    CellBorder.Create("000000", style),
                    CellBorder.Create("000000", style)
                )));
            row++;
        }
        
        sheet.SetColumnWith(0, 20.0);
        sheet.SetColumnWith(1, 20.0);
        
        var workbook = new WorkBook("BorderStyles", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_06_BorderStylesMatrix.xlsx");
    }
}
