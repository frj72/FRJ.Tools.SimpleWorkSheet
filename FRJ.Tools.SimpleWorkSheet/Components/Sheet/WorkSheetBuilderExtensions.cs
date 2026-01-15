using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class WorkSheetBuilderExtensions
{
    extension(WorkSheet sheet)
    {
        public Cell AddEmptyCell(uint x, uint y, Action<CellBuilder> configure)
            => sheet.AddEmptyCell(new(x, y), configure);

        public Cell AddCell(uint x, uint y, CellValue value, Action<CellBuilder>? configure)
            => sheet.AddCell(new(x, y), value, configure);

        public Cell AddStyledCell(CellPosition position, CellValue value, CellStyle style)
            => sheet.AddCell(position, value, cell => cell.WithStyle(style));

        public Cell AddStyledCell(uint x, uint y, CellValue value, CellStyle style)
            => sheet.AddCell(new(x, y), value, cell => cell.WithStyle(style));

        public IEnumerable<Cell> AddRow(uint row, uint startColumn, IEnumerable<CellValue> values,
            Action<CellBuilder>? configure)
        {
            var cells = new List<Cell>();
            var column = startColumn;

            foreach (var value in values)
            {
                var cell = sheet.AddCell(column, row, value, configure: configure);
                cells.Add(cell);
                column++;
            }

            return cells;
        }

        public IEnumerable<Cell> AddColumn(uint column, uint startRow, IEnumerable<CellValue> values,
            Action<CellBuilder>? configure)
        {
            var cells = new List<Cell>();
            var row = startRow;

            foreach (var value in values)
            {
                var cell = sheet.AddCell(column, row, value, configure: configure);
                cells.Add(cell);
                row++;
            }

            return cells;
        }

        public Cell UpdateCell(uint x, uint y, Action<CellBuilder> configure)
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
}
