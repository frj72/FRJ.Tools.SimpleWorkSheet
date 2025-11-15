#r "FRJ.Tools.SimpleWorkSheet/bin/Release/net10.0/FRJ.Tools.SimpleWorkSheet.dll"
#r "FRJ.Tools.SimpleWorkSheet/bin/Release/net10.0/DocumentFormat.OpenXml.dll"
#r "FRJ.Tools.SimpleWorkSheet/bin/Release/net10.0/OneOf.dll"

using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Components.Book;

var sheet = new WorkSheet("FreezePanes");

var headers = new[] { "Column A", "Column B", "Column C", "Column D", "Column E" }
    .Select(h => new CellValue(h));
sheet.AddRow(0, 0, headers, cell => cell
    .WithFont(font => font.Bold())
    .WithColor("4472C4"));

for (uint row = 1; row < 20; row++)
{
    for (uint col = 0; col < 5; col++)
    {
        sheet.AddCell(new CellPosition(col, row), $"R{row}C{col}");
    }
}

sheet.FreezePanes(1, 0);

var workbook = new WorkBook("Test", new[] { sheet });
workbook.SaveToFile("FRJ.Tools.SimpleWorkSheet.Examples/Output/32_FreezePanes.xlsx");
Console.WriteLine("Generated 32_FreezePanes.xlsx");
