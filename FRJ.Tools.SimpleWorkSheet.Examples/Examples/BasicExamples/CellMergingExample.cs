using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;

public class CellMergingExample : IExample
{
    public string Name => "Cell Merging";
    public string Description => "Demonstrates merging cells to create grouped headers";

    public void Run()
    {
        var sheet = new WorkSheet("Cell Merging");

        sheet.AddCell(new CellPosition(0, 0), "Quarterly Sales", cell => cell
            .WithFont(font => font.Bold().WithSize(16))
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Center))
            .WithColor("FFD966"));
        sheet.MergeCells(0, 0, 4, 0);

        sheet.AddCell(new CellPosition(0, 1), "Product", cell => cell.WithFont(font => font.Bold()));
        sheet.MergeCells(0, 1, 0, 2);
        sheet.AddCell(new CellPosition(1, 1), "Q1", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new CellPosition(2, 1), "Q2", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new CellPosition(3, 1), "Q3", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new CellPosition(4, 1), "Q4", cell => cell.WithFont(font => font.Bold()));

        sheet.AddCell(new CellPosition(1, 2), "North", cell => cell.WithFont(font => font.Bold()));
        sheet.MergeCells(1, 2, 2, 2);
        sheet.AddCell(new CellPosition(3, 2), "South", cell => cell.WithFont(font => font.Bold()));
        sheet.MergeCells(3, 2, 4, 2);

        sheet.AddCell(new CellPosition(1, 3), 125000);
        sheet.AddCell(new CellPosition(2, 3), 118500);
        sheet.AddCell(new CellPosition(3, 3), 98000);
        sheet.AddCell(new CellPosition(4, 3), 102500);

        sheet.SetColumnWith(0, 18.0);
        sheet.SetColumnWith(1, 14.0);
        sheet.SetColumnWith(2, 14.0);
        sheet.SetColumnWith(3, 14.0);
        sheet.SetColumnWith(4, 14.0);

        ExampleRunner.SaveWorkSheet(sheet, "37_CellMerging.xlsx");
    }
}
