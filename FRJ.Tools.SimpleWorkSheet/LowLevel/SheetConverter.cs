using System.Globalization;
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
                
            if (workBook.NamedRanges.Count != 0)
            {
                var definedNames = new DefinedNames();
                foreach (var namedRange in workBook.NamedRanges)
                {
                    var definedName = new DefinedName
                    {
                        Name = namedRange.Name,
                        Text = namedRange.ToFormulaReference()
                    };
                    definedNames.Append(definedName);
                }
                workbookPart.Workbook.Append(definedNames);
            }

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
                        if (rowHeight.IsT0)
                        {
                            row.Height = rowHeight.AsT0;
                            row.CustomHeight = true;
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
                {
                    var column = new Column
                    {
                        Min = columnWidth.Key + 1,
                        Max = columnWidth.Key + 1
                    };

                    if (columnWidth.Value.IsT0)
                    {
                        column.Width = columnWidth.Value.AsT0;
                        column.CustomWidth = true;
                    }

                    if (columnWidth.Value is { IsT1: true, AsT1: CellWidth.AutoExpand })
                    {
                        column.BestFit = true;
                        column.CustomWidth = false;
                    }

                    columns.Append(column);
                    
                }

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
                    
                    worksheetPart.Worksheet.InsertAt(sheetViews, 0);
                }

                if (workSheet.MergedCells.Count != 0)
                {
                    var mergeCells = new MergeCells();
                    foreach (var range in workSheet.MergedCells)
                    {
                        var reference = $"{GetCellReference(range.From)}:{GetCellReference(range.To)}";
                        mergeCells.Append(new MergeCell { Reference = reference });
                    }
                    worksheetPart.Worksheet.InsertAfter(mergeCells, sheetData);
                }

                if (columns.Any())
                {
                    var insertIndex = workSheet.FrozenPane != null ? 1 : 0;
                    worksheetPart.Worksheet.InsertAt(columns, insertIndex);
                }

                var cellsWithHyperlinks = workSheet.Cells.Cells.Where(c => c.Value.Hyperlink != null).ToList();

                if (cellsWithHyperlinks.Count != 0)
                {
                    var hyperlinks = new Hyperlinks();
                    
                    foreach (var cellEntry in cellsWithHyperlinks)
                    {
                        var position = cellEntry.Key;
                        var cell = cellEntry.Value;
                        var hyperlinkInfo = cell.Hyperlink;
                        
                        if (hyperlinkInfo == null)
                            continue;
                        
                        var hyperlinkRelationship = worksheetPart.AddHyperlinkRelationship(new(hyperlinkInfo.Url, UriKind.Absolute), true);
                        
                        var hyperlink = new Hyperlink
                        {
                            Reference = GetCellReference(position),
                            Id = hyperlinkRelationship.Id
                        };
                        
                        if (!string.IsNullOrEmpty(hyperlinkInfo.Tooltip))
                            hyperlink.Tooltip = hyperlinkInfo.Tooltip;
                        
                        hyperlinks.Append(hyperlink);
                    }
                    
                    worksheetPart.Worksheet.Append(hyperlinks);
                }

                if (workSheet.Validations.Count != 0)
                {
                    var dataValidations = new DataValidations();
                    
                    foreach (var validationEntry in workSheet.Validations)
                    {
                        var range = validationEntry.Key;
                        var validation = validationEntry.Value;
                        
                        var dataValidation = new DataValidation
                        {
                            Type = ConvertValidationType(validation.Type),
                            Operator = ConvertValidationOperator(validation.Operator),
                            AllowBlank = validation.AllowBlank,
                            ShowInputMessage = validation.ShowInputMessage,
                            ShowErrorMessage = validation.ShowErrorAlert,
                            SequenceOfReferences = new ListValue<StringValue> 
                            { 
                                InnerText = $"{GetCellReference(range.From)}:{GetCellReference(range.To)}" 
                            }
                        };
                        
                        if (!string.IsNullOrEmpty(validation.Formula1))
                        {
                            dataValidation.Formula1 = new Formula1(validation.Formula1);
                        }
                        
                        if (!string.IsNullOrEmpty(validation.Formula2))
                        {
                            dataValidation.Formula2 = new Formula2(validation.Formula2);
                        }
                        
                        if (validation.ShowInputMessage)
                        {
                            if (!string.IsNullOrEmpty(validation.InputTitle))
                                dataValidation.PromptTitle = validation.InputTitle;
                            if (!string.IsNullOrEmpty(validation.InputMessage))
                                dataValidation.Prompt = validation.InputMessage;
                        }
                        
                        if (validation.ShowErrorAlert)
                        {
                            if (!string.IsNullOrEmpty(validation.ErrorTitle))
                                dataValidation.ErrorTitle = validation.ErrorTitle;
                            if (!string.IsNullOrEmpty(validation.ErrorMessage))
                                dataValidation.Error = validation.ErrorMessage;
                            if (!string.IsNullOrEmpty(validation.ErrorStyle))
                                dataValidation.ErrorStyle = ConvertErrorStyle(validation.ErrorStyle);
                        }
                        
                        dataValidations.Append(dataValidation);
                    }
                    
                    dataValidations.Count = (uint)workSheet.Validations.Count;
                    worksheetPart.Worksheet.Append(dataValidations);
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
            newCell.CellValue = new(cellValue.AsDecimal().ToString(CultureInfo.InvariantCulture));
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
            newCell.CellValue = new(oaDate.ToString(CultureInfo.InvariantCulture));
            newCell.DataType = CellValues.Number;
        }
        else if (cellValue.IsDateTimeOffset())
        {
            var dateTimeOffset = cellValue.AsDateTimeOffset();
            var oaDate = dateTimeOffset.DateTime.ToOADate();
            newCell.CellValue = new(oaDate.ToString(CultureInfo.InvariantCulture));
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

    private static DataValidationValues ConvertValidationType(ValidationType type) =>
        type switch
        {
            ValidationType.List => DataValidationValues.List,
            ValidationType.WholeNumber => DataValidationValues.Whole,
            ValidationType.Decimal => DataValidationValues.Decimal,
            ValidationType.Date => DataValidationValues.Date,
            ValidationType.Time => DataValidationValues.Time,
            ValidationType.TextLength => DataValidationValues.TextLength,
            ValidationType.Custom => DataValidationValues.Custom,
            _ => DataValidationValues.None
        };

    private static DataValidationOperatorValues ConvertValidationOperator(ValidationOperator op) =>
        op switch
        {
            ValidationOperator.Between => DataValidationOperatorValues.Between,
            ValidationOperator.NotBetween => DataValidationOperatorValues.NotBetween,
            ValidationOperator.Equal => DataValidationOperatorValues.Equal,
            ValidationOperator.NotEqual => DataValidationOperatorValues.NotEqual,
            ValidationOperator.GreaterThan => DataValidationOperatorValues.GreaterThan,
            ValidationOperator.LessThan => DataValidationOperatorValues.LessThan,
            ValidationOperator.GreaterThanOrEqual => DataValidationOperatorValues.GreaterThanOrEqual,
            ValidationOperator.LessThanOrEqual => DataValidationOperatorValues.LessThanOrEqual,
            _ => DataValidationOperatorValues.Between
        };

    private static DataValidationErrorStyleValues ConvertErrorStyle(string style) =>
        style.ToLowerInvariant() switch
        {
            "stop" => DataValidationErrorStyleValues.Stop,
            "warning" => DataValidationErrorStyleValues.Warning,
            "information" => DataValidationErrorStyleValues.Information,
            _ => DataValidationErrorStyleValues.Stop
        };
}


