using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class WorkSheetBuilderExtensions
{
    public static Cell AddCell(this WorkSheet sheet, uint x, uint y, Action<CellBuilder> configure)
    {
        return sheet.AddCell(new CellPosition(x, y), configure);
    }

    public static Cell AddCell(this WorkSheet sheet, uint x, uint y, CellValue value, Action<CellBuilder>? configure = null)
    {
        return sheet.AddCell(new CellPosition(x, y), value, configure);
    }

    public static Cell AddStyledCell(this WorkSheet sheet, CellPosition position, CellValue value, CellStyle style)
    {
        return sheet.AddCell(position, value, cell => cell.WithStyle(style));
    }

    public static Cell AddStyledCell(this WorkSheet sheet, uint x, uint y, CellValue value, CellStyle style)
    {
        return sheet.AddCell(new CellPosition(x, y), value, cell => cell.WithStyle(style));
    }

    public static IEnumerable<Cell> AddRow(this WorkSheet sheet, uint row, uint startColumn, IEnumerable<CellValue> values, Action<CellBuilder>? configure = null)
    {
        var cells = new List<Cell>();
        var column = startColumn;
        
        foreach (var value in values)
        {
            var cell = sheet.AddCell(column, row, value, configure);
            cells.Add(cell);
            column++;
        }
        
        return cells;
    }

    public static IEnumerable<Cell> AddColumn(this WorkSheet sheet, uint column, uint startRow, IEnumerable<CellValue> values, Action<CellBuilder>? configure = null)
    {
        var cells = new List<Cell>();
        var row = startRow;
        
        foreach (var value in values)
        {
            var cell = sheet.AddCell(column, row, value, configure);
            cells.Add(cell);
            row++;
        }
        
        return cells;
    }

    public static Cell UpdateCell(this WorkSheet sheet, uint x, uint y, Action<CellBuilder> configure)
    {
        var position = new CellPosition(x, y);
        var existingCell = sheet.Cells.Cells.TryGetValue(position, out var cell) ? cell : Cell.CreateEmpty();
        
        var builder = CellBuilder.FromCell(existingCell);
        configure(builder);
        var updatedCell = builder.Build();
        
        if (!updatedCell.Style.HasValidColors())
            throw new ArgumentException("Invalid cell style colors", nameof(configure));
        
        sheet.Cells.Cells[position] = updatedCell;
        return updatedCell;
    }
}
