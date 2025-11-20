using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using OneOf;
using ArgumentException = System.ArgumentException;
// ReSharper disable UnusedMember.Global

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public class WorkSheet
{


    private readonly List<CellRange> _mergedCells = [];
    private readonly List<Chart> _charts = [];
    private readonly List<ExcelTable> _tables = [];
    private readonly List<WorksheetImage> _images = [];
    public string Name { get; }
    public CellCollection Cells { get; }
    
    public Dictionary<uint, OneOf<double, CellWidth>> ExplicitColumnWidths { get; } = new();
    public Dictionary<uint, OneOf<double, RowHeight>> ExplicitRowHeights { get; } = new();
    public Dictionary<CellRange, CellValidation> Validations { get; } = new();
    public FreezePane? FrozenPane { get; private set; }
    public string? TabColor { get; private set; }
    public bool IsVisible { get; private set; } = true;
    public IReadOnlyList<CellRange> MergedCells => _mergedCells;
    public IReadOnlyList<Chart> Charts => _charts;
    public IReadOnlyList<ExcelTable> Tables => _tables;
    public IReadOnlyList<WorksheetImage> Images => _images;

    public WorkSheet(string name)
    {
        Name = name;
        Cells = new();
    }

    public Cell AddEmptyCell(CellPosition position, Action<CellBuilder> configure)
    {
        var builder = CellBuilder.Create();
        configure(builder);
        var cell = builder.Build();

        if (!cell.Style.HasValidColors())
            throw new ArgumentException("Invalid cell style colors", nameof(configure));

        Cells.Cells[position] = cell;
        return cell;
    }

    public Cell AddCell(CellPosition position, CellValue value, Action<CellBuilder>? configure)
    {
        var builder = CellBuilder.FromValue(value);
        configure?.Invoke(builder);
        var cell = builder.Build();

        if (!cell.Style.HasValidColors())
            throw new ArgumentException("Invalid cell style colors");

        Cells.Cells[position] = cell;
        return cell;
    }



    private void AddCellLegacy(CellPosition position, CellValue value, string? color = null, CellFont? font = null,
        CellBorders? borders = null)
    {
        if (!color.IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(color));
        if (!font.HasValidColor())
            throw new ArgumentException("Invalid font color format", nameof(font));
        if (!borders.HasValidColors())
            throw new ArgumentException("Invalid border color format", nameof(borders));

        var cell = Cell.Create(value, color ?? WorkSheetDefaults.FillColor, font ?? WorkSheetDefaults.Font,
            borders ?? WorkSheetDefaults.CellBorders, null);
        Cells.Cells[position] = cell;
    }

    public CellValue? GetValue(uint x, uint y)
    {
        var position = new CellPosition(x, y);
        return Cells.Cells.TryGetValue(position, out var cell) ? cell.Value : null;
    }

    public void InsertCell(int x, int y, Cell cell) => Cells.Cells[new((uint)x, (uint)y)] = cell;


    public void SetValue(uint x, uint y, CellValue value)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
            AddCellLegacy(new(x, y), value);
        else
            Cells.Cells[position] = cell.SetValue(value);
        
    }

    public void SetFont(uint x, uint y, CellFont font)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
             AddCellLegacy(new(x, y), string.Empty, font: font);
        else
            Cells.Cells[position] = cell.SetFont(font);
    }
    

    public void SetColor(uint x, uint y, string color)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
            AddCellLegacy(new(x, y), string.Empty, color: color);
        else
            Cells.Cells[position] = cell.SetColor(color);
    }
    

    public void SetBorders(uint x, uint y, CellBorders borders)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
             AddCellLegacy(new(x, y), string.Empty, borders: borders);
        else
            Cells.Cells[position] = cell.SetBorders(borders);
    }
    

    public void SetBorders(uint x, uint y, CellBorder borders) 
        => SetBorders(x, y, CellBorders.Create(borders, borders, borders, borders));

    public void SetDefaultFormatting(uint x, uint y)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
            AddCellLegacy(new(x, y), string.Empty);
        else
        {
            cell = cell.SetDefaultFormatting();
            Cells.Cells[position] = cell;
        }
    }
    
    public  void SetColumnWidth(uint column, OneOf<double, CellWidth> width) 
        => ExplicitColumnWidths[column] = width;

    public void SetRowHeight(uint row, OneOf<double, RowHeight> height) 
        => ExplicitRowHeights[row] = height;

    public void FreezePanes(uint row, uint column) => FrozenPane = new(row, column);

    public void FreezeRows(uint row) => FrozenPane = new(row, 0);

    public void FreezeColumns(uint column) => FrozenPane = new(0, column);

    public CellRange MergeCells(uint fromX, uint fromY, uint toX, uint toY) =>
        MergeCells(CellRange.FromBounds(fromX, fromY, toX, toY));

    public CellRange MergeCells(CellPosition from, CellPosition to) =>
        MergeCells(CellRange.FromPositions(from, to));

    public CellRange MergeCells(CellRange range)
    {
        if (range.IsSingleCell)
            throw new ArgumentException("Merged range must span at least two cells", nameof(range));

        if (_mergedCells.Any(existing => existing.Overlaps(range)))
            throw new ArgumentException("Merged range overlaps an existing merged range", nameof(range));

        EnsureCellExists(range.From);
        _mergedCells.Add(range);
        return range;
    }

    internal void ImportMergedRange(CellRange range)
    {
        if (range.IsSingleCell)
            return;

        EnsureCellExists(range.From);

        if (_mergedCells.Any(existing => existing.Overlaps(range)))
            return;

        _mergedCells.Add(range);
    }

    public HashSet<int> GetAllFontSizes() => 
        Cells.Cells.Values.Select(cell => cell.Font?.Size ?? WorkSheetDefaults.FontSize).ToHashSet();

    public HashSet<string> GetAllFontNames() => 
        Cells.Cells.Values.Select(cell => cell.Font?.Name ?? WorkSheetDefaults.FontName).ToHashSet();

    public HashSet<string> GetAllColors()
    {
        var cellColors = Cells.Cells.Values.Select(cell => cell.Color ?? WorkSheetDefaults.Color);
        var fontColors = Cells.Cells.Values.Select(cell => cell.Font?.Color ?? WorkSheetDefaults.Color);
        var borderColors = Cells.Cells.Values.SelectMany(cell => cell.Borders?.GetAllColors() ?? [WorkSheetDefaults.Color
        ]);
        return cellColors.Concat(fontColors).Concat(borderColors).ToHashSet();
    }

    public HashSet<string> GetAllCellColors() 
        => Cells.Cells.Values.Select(cell => cell.Color ?? WorkSheetDefaults.Color).ToHashSet();


    public HashSet<CellFont> GetAllFonts() 
        => Cells.Cells.Values.Select(cell => cell.Font ?? WorkSheetDefaults.Font).ToHashSet();

    public void AddValidation(uint x, uint y, CellValidation validation)
    {
        var position = new CellPosition(x, y);
        var range = CellRange.FromPositions(position, position);
        Validations[range] = validation;
    }

    public void AddValidation(CellPosition position, CellValidation validation)
    {
        var range = CellRange.FromPositions(position, position);
        Validations[range] = validation;
    }

    public void AddValidation(CellRange range, CellValidation validation) 
        => Validations[range] = validation;

    public void AddValidation(uint fromX, uint fromY, uint toX, uint toY, CellValidation validation)
    {
        var range = CellRange.FromBounds(fromX, fromY, toX, toY);
        Validations[range] = validation;
    }
    
    public void AddChart(Chart chart)
    {
        if (chart.Position == null)
            throw new ArgumentException("Chart must have a position before being added to worksheet", nameof(chart));

        _charts.Add(chart);
    }

    public void AutoFitColumn(uint column)
    {
        var cellsInColumn = Cells.Cells
            .Where(kvp => kvp.Key.X == column)
            .Select(kvp => kvp.Value);

        var estimatedWidth = cellsInColumn.EstimateMaxWidth();
        SetColumnWidth(column, estimatedWidth);
    }

    public void AutoFitAllColumns()
    {
        var columns = Cells.Cells.Keys.Select(pos => pos.X).Distinct();
        foreach (var column in columns)
            AutoFitColumn(column);
    }

    public void SetTabColor(string color)
    {
        if (!color.IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(color));
        TabColor = color;
    }

    public void SetVisible(bool visible) => IsVisible = visible;

    public ExcelTable AddTable(string name, CellRange range, bool showFilterButton = true)
    {
        if (_tables.Any(t => t.Name == name))
            throw new ArgumentException($"Table with name '{name}' already exists", nameof(name));
        
        if (_tables.Any(t => t.Range.Overlaps(range)))
            throw new ArgumentException("Table range overlaps an existing table", nameof(range));
        
        var table = new ExcelTable(name, range, showFilterButton, null);
        _tables.Add(table);
        return table;
    }

    public ExcelTable AddTable(string name, uint fromX, uint fromY, uint toX, uint toY, bool showFilterButton = true)
    {
        var range = CellRange.FromBounds(fromX, fromY, toX, toY);
        return AddTable(name, range, showFilterButton);
    }

    public void AddImage(WorksheetImage image) 
        => _images.Add(image);

    public void AddImage(string filePath, CellPosition position, uint widthInPixels, uint heightInPixels)
    {
        var image = WorksheetImage.FromFile(filePath, position, widthInPixels, heightInPixels);
        AddImage(image);
    }
    
    private void EnsureCellExists(CellPosition position)
    {
        if (Cells.Cells.ContainsKey(position))
            return;
        AddCell(position, string.Empty, null);
    }
}

