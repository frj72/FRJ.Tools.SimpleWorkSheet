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

        sheet.AddCell(new(0, 0), "Quarterly Sales", cell => cell
            .WithFont(font => font.Bold().WithSize(16))
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Center))
            .WithColor("FFD966"));
        sheet.MergeCells(0, 0, 4, 0);

        sheet.AddCell(new(0, 1), "Product", cell => cell.WithFont(font => font.Bold()));
        sheet.MergeCells(0, 1, 0, 2);
        sheet.AddCell(new(1, 1), "Q1", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(2, 1), "Q2", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(3, 1), "Q3", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(4, 1), "Q4", cell => cell.WithFont(font => font.Bold()));

        sheet.AddCell(new(1, 2), "North", cell => cell.WithFont(font => font.Bold()));
        sheet.MergeCells(1, 2, 2, 2);
        sheet.AddCell(new(3, 2), "South", cell => cell.WithFont(font => font.Bold()));
        sheet.MergeCells(3, 2, 4, 2);

        sheet.AddCell(new(1, 3), 125000, null);
        sheet.AddCell(new(2, 3), 118500, null);
        sheet.AddCell(new(3, 3), 98000, null);
        sheet.AddCell(new(4, 3), 102500, null);

        sheet.SetColumnWidth(0, 18.0);
        sheet.SetColumnWidth(1, 14.0);
        sheet.SetColumnWidth(2, 14.0);
        sheet.SetColumnWidth(3, 14.0);
        sheet.SetColumnWidth(4, 14.0);

        ExampleRunner.SaveWorkSheet(sheet, "37_CellMerging.xlsx");
    }
}
