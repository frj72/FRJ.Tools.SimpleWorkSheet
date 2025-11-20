using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorkSheet.Components.Book;

public class WorkBook
{
    private readonly List<NamedRange> _namedRanges = [];

    public WorkBook(string name, IEnumerable<WorkSheet> sheets)
    {
        Name = name;
        Sheets = sheets;
    }

    public string Name { get; }
    public IEnumerable<WorkSheet> Sheets { get; }

    public IReadOnlyList<NamedRange> NamedRanges => _namedRanges;

    public void SaveToFile(string fileName)
    {
        var bytes = SheetConverter.ToBinaryExcelFile(this);
        File.WriteAllBytes(fileName, bytes);
    }

    public void AddNamedRange(string name, string sheetName, CellRange range)
    {
        if (NamedRanges.Any(nr => nr.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Named range '{name}' already exists", nameof(name));

        var namedRange = new NamedRange(name, sheetName, range);
        _namedRanges.Add(namedRange);
    }

    public void AddNamedRange(string name, string sheetName, uint fromX, uint fromY, uint toX, uint toY)
    {
        var range = CellRange.FromBounds(fromX, fromY, toX, toY);
        AddNamedRange(name, sheetName, range);
    }
}
