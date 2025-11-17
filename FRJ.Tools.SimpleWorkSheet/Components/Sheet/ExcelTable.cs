namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public record ExcelTable
{
    public string Name { get; init; }
    public CellRange Range { get; init; }
    public bool ShowFilterButton { get; init; }
    public string? DisplayName { get; init; }

    public ExcelTable(string name, CellRange range, bool showFilterButton = true, string? displayName = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Table name cannot be empty", nameof(name));
        
        if (range.IsSingleCell)
            throw new ArgumentException("Table range must span at least two cells", nameof(range));
        
        Name = name;
        Range = range;
        ShowFilterButton = showFilterButton;
        DisplayName = displayName ?? name;
    }
}
