using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
// ReSharper disable UnusedMember.Global

namespace FRJ.Tools.SimpleWorkSheet.LowLevel;

public class SheetConverter
{
    public static byte[] ToBinaryExcelFile(WorkSheet workSheet)
    {
        var workbook = new WorkBook("Workbook", [workSheet]);
        return ToBinaryExcelFile(workbook);
    }
    public static byte[] ToBinaryExcelFile(WorkBook workBook)
    {
        using var memoryStream = new MemoryStream();
        using (var spreadsheetDocument =
               SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new();
            
            workbookPart.Workbook.Append(new BookViews(new WorkbookView()));
            
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var styleHelper = new StyleHelper();
            styleHelper.CollectStyles(workBook);
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = styleHelper.GenerateStylesheet();
            stylesPart.Stylesheet.Save();
            uint sheetId = 1;
                
            foreach (var workSheet in workBook.Sheets)
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new(new SheetData());
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() ?? throw new InvalidOperationException();

               
                var rows = workSheet.Cells.Cells.GroupBy(cellEntry => cellEntry.Key.Y)
                    .OrderBy(g => g.Key);

                foreach (var rowGroup in rows)
                {
                    var row = new Row { RowIndex = rowGroup.Key + 1 };

                    if (workSheet.ExplicitRowHeights.TryGetValue(rowGroup.Key, out var rowHeight))
                    {
                        if (rowHeight.IsT0)
                        {
                            row.Height = rowHeight.AsT0;
                            row.CustomHeight = true;
                        }
                    }

                    foreach (var cellEntry in rowGroup.OrderBy(ce => ce.Key.X))
                    {
                        var position = cellEntry.Key;
                        var cell = cellEntry.Value;
                            
                        var newCell = new Cell
                        {
                            CellReference = GetCellReference(position)
                        };
                            
                        AssignCellValue(newCell, cell);
                            
                        var styleIndex = styleHelper.GetStyleIndex(cell);
                        newCell.StyleIndex = styleIndex;

                        row.Append(newCell);
                    }

                    sheetData.Append(row);
                }
                
                var columns = new Columns();
                foreach (var columnWidth in workSheet.ExplicitColumnWidths)
                    if (columnWidth.Value.IsT0)
                    {
                        var column = new Column
                        {
                            Min = columnWidth.Key + 1,
                            Max = columnWidth.Key + 1,
                            Width = columnWidth.Value.AsT0,
                            CustomWidth = true
                        };
                        columns.Append(column);
                    }
                    else if (columnWidth.Value is { IsT1: true, AsT1: CellWidth.AutoExpand })
                    {
                        var cellsInColumn = workSheet.Cells.Cells.Where(x=> x.Key.X == columnWidth.Key).Select(x=> x.Value).ToList();
                        if (cellsInColumn.Count == 0)
                            continue;
                        var maxCellWidth = cellsInColumn.EstimateMaxWidth();
                        var column = new Column
                        {
                            Min = columnWidth.Key + 1,
                            Max = columnWidth.Key + 1,
                            Width = maxCellWidth,
                            CustomWidth = true
                        };
                        columns.Append(column);
                        
                    }

                if (columns.Any())
                    worksheetPart.Worksheet.InsertAt(columns, 0);

                if (workSheet.FrozenPane != null)
                {
                    var sheetViews = new SheetViews();
                    var sheetView = new SheetView { WorkbookViewId = 0 };
                    
                    var pane = new Pane
                    {
                        TopLeftCell = GetCellReference(new(workSheet.FrozenPane.Column, workSheet.FrozenPane.Row)),
                        State = PaneStateValues.Frozen
                    };
                    
                    if (workSheet.FrozenPane.Column > 0)
                        pane.HorizontalSplit = (double)workSheet.FrozenPane.Column;
                    
                    if (workSheet.FrozenPane.Row > 0)
                        pane.VerticalSplit = (double)workSheet.FrozenPane.Row;
                    
                    var activePane = (workSheet.FrozenPane.Row > 0, workSheet.FrozenPane.Column > 0) switch
                    {
                        (true, true) => PaneValues.BottomRight,
                        (true, false) => PaneValues.BottomLeft,
                        (false, true) => PaneValues.TopRight,
                        _ => PaneValues.TopLeft
                    };
                    pane.ActivePane = activePane;
                    
                    sheetView.Append(pane);
                    
                    var selection = new Selection { Pane = activePane };
                    sheetView.Append(selection);
                    
                    sheetViews.Append(sheetView);
                    
                    worksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);
                }

                worksheetPart.Worksheet.Save();
                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId++,
                    Name = workSheet.Name
                };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();
            }

            workbookPart.Workbook.Save();
        }

        return memoryStream.ToArray();
    }

    private static string GetCellReference(CellPosition position)
    {
        var columnName = GetExcelColumnName(position.X + 1);
        var cellReference = $"{columnName}{position.Y + 1}";
        return cellReference;
    }

    private static string GetExcelColumnName(uint columnNumber)
    {
        var dividend = columnNumber;
        var columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    private static void AssignCellValue(Cell newCell, Components.SimpleCell.Cell cell)
    {
        var cellValue = cell.Value;

        if (cellValue.IsDecimal())
        {
            newCell.DataType = CellValues.Number;
            newCell.CellValue = new(cellValue.AsDecimal().ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        else if (cellValue.IsLong())
        {
            newCell.DataType = CellValues.Number;
            newCell.CellValue = new(cellValue.AsLong().ToString());
        }
        else if (cellValue.IsDateTime())
        {
            var dateTime = cellValue.AsDateTime();
            var oaDate = dateTime.ToOADate();
            newCell.CellValue = new(oaDate.ToString(System.Globalization.CultureInfo.InvariantCulture));
            newCell.DataType = CellValues.Number;
        }
        else if (cellValue.IsDateTimeOffset())
        {
            var dateTimeOffset = cellValue.AsDateTimeOffset();
            var oaDate = dateTimeOffset.DateTime.ToOADate();
            newCell.CellValue = new(oaDate.ToString(System.Globalization.CultureInfo.InvariantCulture));
            newCell.DataType = CellValues.Number;
        }
        else if (cellValue.IsString())
        {
            newCell.DataType = CellValues.String;
            newCell.CellValue = new(cellValue.AsString());
        }
        else if (cellValue.IsCellFormula())
            newCell.CellFormula = new(cellValue.Value.AsT5.Value);
        else
        {
            newCell.DataType = CellValues.String;
            newCell.CellValue = new(cellValue.AsString());
        }
    }
}


