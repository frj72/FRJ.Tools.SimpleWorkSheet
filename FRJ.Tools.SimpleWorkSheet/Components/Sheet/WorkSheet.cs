using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using OneOf;
using ArgumentException = System.ArgumentException;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public class WorkSheet
{


    public string Name { get; }
    public CellCollection Cells { get; }
    
    public Dictionary<uint, OneOf.OneOf<double, CellWidth>> ExplicitColumnWidths { get; } = new();

    public WorkSheet(string name)
    {
        Name = name;
        Cells = new();
    }

    public Cell AddCell(CellPosition position, CellValue value, string? color = null, CellFont? font = null,
        CellBorders? borders = null)
    {
        if (!color.IsValidColor())
            throw new ArgumentException("Invalid color format", nameof(color));
        if (!font.HasValidColor())
            throw new ArgumentException("Invalid font color format", nameof(font));
        if (!borders.HasValidColors())
            throw new ArgumentException("Invalid border color format", nameof(borders));

        var cell = Cell.Create(value, color ?? WorkSheetDefaults.FillColor, font ?? WorkSheetDefaults.Font,
            borders ?? WorkSheetDefaults.CellBorders);
        Cells.Cells[position] = cell;
        return cell;
    }

    public CellValue? GetValue(uint x, uint y)
    {
        var position = new CellPosition(x, y);
        return Cells.Cells.TryGetValue(position, out var cell) ? cell.Value : null;
    }

    public void InsertCell(int x, int y, Cell cell)
    {
        Cells.Cells[new CellPosition((uint)x, (uint)y)] = cell;
    }


    public void SetValue(uint x, uint y, CellValue value)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
            AddCell(new(x, y), value);
        else
            Cells.Cells[position] = cell.SetValue(value);
        
    }

    public void SetFont(uint x, uint y, CellFont font)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
             AddCell(new(x, y), string.Empty, font: font);
        else
            Cells.Cells[position] = cell.SetFont(font);
    }
    

    public void SetColor(uint x, uint y, string color)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
            AddCell(new(x, y), string.Empty, color: color);
        else
            Cells.Cells[position] = cell.SetColor(color);
    }
    


    public void SetBorders(uint x, uint y, CellBorders borders)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
             AddCell(new(x, y), string.Empty, borders: borders);
        else
            Cells.Cells[position] = cell.SetBorders(borders);
    }
    

    public void SetBorders(uint x, uint y, CellBorder borders)
    {
        SetBorders(x, y, CellBorders.Create(borders, borders, borders, borders));
    }

    public void SetDefaultFormatting(uint x, uint y)
    {
        var position = new CellPosition(x, y);
        if (!Cells.Cells.TryGetValue(position, out var cell)) 
            AddCell(new(x, y), string.Empty);
        else
        {
            cell = cell.SetDefaultFormatting();
            Cells.Cells[position] = cell;
        }
    }
    
    public  void SetColumnWith(uint column, OneOf<double, CellWidth> width)
    {
        ExplicitColumnWidths[column] = width;
    }

    public HashSet<int> GetAllFontSizes()
    {
        return Cells.Cells.Values.Select(cell => cell.Font?.Size ?? WorkSheetDefaults.FontSize).ToHashSet();
    }

    public HashSet<string> GetAllFontNames()
    {
        return Cells.Cells.Values.Select(cell => cell.Font?.Name ?? WorkSheetDefaults.FontName).ToHashSet();
    }

    public HashSet<string> GetAllColors()
    {
        var cellColors = Cells.Cells.Values.Select(cell => cell.Color ?? WorkSheetDefaults.Color);
        var fontColors = Cells.Cells.Values.Select(cell => cell.Font?.Color ?? WorkSheetDefaults.Color);
        var borderColors = Cells.Cells.Values.SelectMany(cell => cell.Borders?.GetAllColors() ?? new[] { WorkSheetDefaults.Color });
        return cellColors.Concat(fontColors).Concat(borderColors).ToHashSet();
    }

    public HashSet<string> GetAllCellColors()
    {
        return Cells.Cells.Values.Select(cell => cell.Color ?? WorkSheetDefaults.Color).ToHashSet();
    }
    

    public HashSet<CellFont> GetAllFonts()
    {
        return Cells.Cells.Values.Select(cell => cell.Font ?? WorkSheetDefaults.Font).ToHashSet();
    }
    
}
