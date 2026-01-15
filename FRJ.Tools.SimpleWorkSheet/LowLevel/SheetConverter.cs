using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using AreaChart = FRJ.Tools.SimpleWorkSheet.Components.Charts.AreaChart;
using BarChart = FRJ.Tools.SimpleWorkSheet.Components.Charts.BarChart;
using BlipFill = DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Chart = FRJ.Tools.SimpleWorkSheet.Components.Charts.Chart;
using Formula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using FromMarker = DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker;
using GraphicFrame = DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame;
using HeaderFooter = DocumentFormat.OpenXml.Drawing.Charts.HeaderFooter;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;
using Index = DocumentFormat.OpenXml.Drawing.Charts.Index;
using LineChart = FRJ.Tools.SimpleWorkSheet.Components.Charts.LineChart;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties;
using NonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties;
using NonVisualGraphicFrameProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties;
using NonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties;
using NonVisualPictureProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties;
using NumberingFormat = DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat;
using OrientationValues = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues;
using PageMargins = DocumentFormat.OpenXml.Drawing.Charts.PageMargins;
using Picture = DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture;
using PieChart = FRJ.Tools.SimpleWorkSheet.Components.Charts.PieChart;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using RunProperties = DocumentFormat.OpenXml.Drawing.RunProperties;
using ScatterChart = FRJ.Tools.SimpleWorkSheet.Components.Charts.ScatterChart;
using Selection = DocumentFormat.OpenXml.Spreadsheet.Selection;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using ToMarker = DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker;
using Values = DocumentFormat.OpenXml.Drawing.Charts.Values;

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
                foreach (var definedName in workBook.NamedRanges.Select(namedRange => new DefinedName
                         {
                             Name = namedRange.Name,
                             Text = namedRange.ToFormulaReference()
                         }))
                    definedNames.Append(definedName);
                workbookPart.Workbook.Append(definedNames);
            }

            foreach (var workSheet in workBook.Sheets)
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new(new SheetData());
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() ?? throw new InvalidOperationException();

                if (!string.IsNullOrEmpty(workSheet.TabColor))
                {
                    var sheetProperties = new SheetProperties(new TabColor { Rgb = workSheet.TabColor });
                    worksheetPart.Worksheet.InsertAt(sheetProperties, 0);
                }
               
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

                        if (rowHeight is { IsT1: true, AsT1: RowHeight.Hidden })
                        {
                            row.Hidden = true;
                            row.CustomHeight = true;
                        }
                    }

                    if (workSheet.HiddenRows.TryGetValue(rowGroup.Key, out var isHidden) && isHidden)
                    {
                        row.Hidden = true;
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

                var allColumnIndices = workSheet.ExplicitColumnWidths.Keys
                    .Concat(workSheet.HiddenColumns.Keys)
                    .Distinct()
                    .OrderBy(x => x);

                foreach (var columnIndex in allColumnIndices)
                {
                    var column = new Column
                    {
                        Min = columnIndex + 1,
                        Max = columnIndex + 1
                    };

                    if (workSheet.ExplicitColumnWidths.TryGetValue(columnIndex, out var columnWidth))
                    {
                        if (columnWidth.IsT0)
                        {
                            column.Width = columnWidth.AsT0;
                            column.CustomWidth = true;
                        }

                        if (columnWidth is { IsT1: true, AsT1: CellWidth.AutoExpand })
                        {
                            column.BestFit = true;
                            column.CustomWidth = false;
                        }

                        if (columnWidth is { IsT1: true, AsT1: CellWidth.Hidden })
                        {
                            column.Hidden = true;
                            column.CustomWidth = true;
                        }
                    }

                    if (workSheet.HiddenColumns.TryGetValue(columnIndex, out var isHidden) && isHidden)
                    {
                        column.Hidden = true;
                        column.CustomWidth = true;
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
                            SequenceOfReferences = new()
                            { 
                                InnerText = $"{GetCellReference(range.From)}:{GetCellReference(range.To)}" 
                            }
                        };
                        
                        if (!string.IsNullOrEmpty(validation.Formula1)) dataValidation.Formula1 = new(validation.Formula1);

                        if (!string.IsNullOrEmpty(validation.Formula2)) dataValidation.Formula2 = new(validation.Formula2);

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

                if (workSheet.Charts.Count != 0 || workSheet.Images.Count != 0)
                {
                    var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                    var drawingId = worksheetPart.GetIdOfPart(drawingsPart);
                    
                    var drawing = new Drawing { Id = drawingId };
                    worksheetPart.Worksheet.Append(drawing);
                    
                    var worksheetDrawing = new WorksheetDrawing();
                    
                    uint chartIndex = 1;
                    foreach (var chart in workSheet.Charts)
                    {
                        var chartPart = drawingsPart.AddNewPart<ChartPart>();
                        var chartRelId = drawingsPart.GetIdOfPart(chartPart);
                        
                        var dataSheetName = chart.DataSourceSheet ?? workSheet.Name;
                        GenerateChartPart(chartPart, chart, dataSheetName);
                        
                        var twoCellAnchor = CreateTwoCellAnchor(chart, chartRelId, chartIndex++);
                        worksheetDrawing.Append(twoCellAnchor);
                    }
                    
                    uint imageIndex = 1;
                    foreach (var image in workSheet.Images)
                    {
                        var imagePartType = image.Format == ImageFormat.Png 
                            ? ImagePartType.Png 
                            : ImagePartType.Jpeg;
                        var imagePart = drawingsPart.AddImagePart(imagePartType);
                        using var imageStream = new MemoryStream(image.ImageData);
                        imagePart.FeedData(imageStream);
                        
                        var imageRelId = drawingsPart.GetIdOfPart(imagePart);
                        var imageTwoCellAnchor = CreateImageTwoCellAnchor(image, imageRelId, imageIndex++);
                        worksheetDrawing.Append(imageTwoCellAnchor);
                    }
                    
                    drawingsPart.WorksheetDrawing = worksheetDrawing;
                    drawingsPart.WorksheetDrawing.Save();
                }

                if (workSheet.Tables.Count != 0)
                {
                    uint tableId = 1;
                    foreach (var excelTable in workSheet.Tables)
                    {
                        var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
                        var tablePartId = worksheetPart.GetIdOfPart(tablePart);
                        
                        var table = new Table
                        {
                            Id = tableId,
                            Name = excelTable.Name,
                            DisplayName = excelTable.DisplayName,
                            Reference = $"{GetCellReference(excelTable.Range.From)}:{GetCellReference(excelTable.Range.To)}",
                            TotalsRowShown = false
                        };
                        
                        if (excelTable.ShowFilterButton)
                        {
                            var autoFilter = new AutoFilter
                            {
                                Reference = $"{GetCellReference(excelTable.Range.From)}:{GetCellReference(excelTable.Range.To)}"
                            };
                            table.Append(autoFilter);
                        }
                        
                        var tableColumns = new TableColumns { Count = excelTable.Range.To.X - excelTable.Range.From.X + 1 };
                        for (uint colIndex = 0; colIndex < tableColumns.Count; colIndex++)
                        {
                            var headerPosition = excelTable.Range.From with { X = excelTable.Range.From.X + colIndex };
                            var columnName = $"Column{colIndex + 1}";
                            
                            if (workSheet.Cells.Cells.TryGetValue(headerPosition, out var headerCell))
                            {
                                var index = colIndex;
                                columnName = headerCell.Value.Value.Match(
                                    d => d.ToString(CultureInfo.InvariantCulture),
                                    l => l.ToString(),
                                    s => s,
                                    dt => dt.ToString(CultureInfo.InvariantCulture),
                                    dto => dto.ToString(),
                                    _ => $"Column{index + 1}"
                                );
                            }

                            var tableColumn = new TableColumn
                            {
                                Id = colIndex + 1,
                                Name = columnName
                            };
                            tableColumns.Append(tableColumn);
                        }
                        table.Append(tableColumns);
                        
                        var tableStyleInfo = new TableStyleInfo
                        {
                            Name = "TableStyleMedium2",
                            ShowFirstColumn = false,
                            ShowLastColumn = false,
                            ShowRowStripes = true,
                            ShowColumnStripes = false
                        };
                        table.Append(tableStyleInfo);
                        
                        tablePart.Table = table;
                        tablePart.Table.Save();
                        
                        var tableParts = worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
                        if (tableParts == null)
                        {
                            tableParts = new();
                            worksheetPart.Worksheet.Append(tableParts);
                        }
                        
                        tableParts.Append(new TablePart { Id = tablePartId });
                        tableParts.Count = (uint)workSheet.Tables.Count;
                        
                        tableId++;
                    }
                }

                worksheetPart.Worksheet.Save();
                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId++,
                    Name = workSheet.Name
                };
                
                if (!workSheet.IsVisible)
                    sheet.State = SheetStateValues.Hidden;
                
                sheets.Append(sheet);
                workbookPart.Workbook.Save();
            }

            var calculationProperties = new CalculationProperties { FullCalculationOnLoad = true };
            workbookPart.Workbook.Append(calculationProperties);

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

    private static void GenerateChartPart(ChartPart chartPart, Chart chart, string sheetName)
    {
        switch (chart)
        {
            case BarChart barChart:
                GenerateBarChartPart(chartPart, barChart, sheetName, chart);
                break;
            case LineChart lineChart:
                GenerateLineChartPart(chartPart, lineChart, sheetName, chart);
                break;
            case PieChart pieChart:
                GeneratePieChartPart(chartPart, pieChart, sheetName, chart);
                break;
            case ScatterChart scatterChart:
                GenerateScatterChartPart(chartPart, scatterChart, sheetName, chart);
                break;
            case AreaChart areaChart:
                GenerateAreaChartPart(chartPart, areaChart, sheetName, chart);
                break;
        }
    }

    private static void AddSeriesColor(OpenXmlCompositeElement seriesElement, string? color, bool includeFill, bool includeLine)
    {
        if (string.IsNullOrEmpty(color))
            return;

        var rgb = color.ToArgbColor()[2..];
        var shapeProperties = new DocumentFormat.OpenXml.Drawing.Charts.ShapeProperties();

        if (includeFill)
        {
            var solidFill = new SolidFill();
            solidFill.Append(new RgbColorModelHex { Val = rgb });
            shapeProperties.Append(solidFill);
        }

        if (includeLine)
        {
            var outline = new DocumentFormat.OpenXml.Drawing.Outline();
            var lineFill = new SolidFill();
            lineFill.Append(new RgbColorModelHex { Val = rgb });
            outline.Append(lineFill);
            shapeProperties.Append(outline);
        }

        seriesElement.Append(shapeProperties);
    }

    private static void GenerateBarChartPart(ChartPart chartPart, BarChart barChart, string sheetName, Chart chart)
    {

        var chartSpace = new ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var editingLanguage = new EditingLanguage { Val = "en-US" };
        chartSpace.Append(editingLanguage);

        var chartElement = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        
        var plotArea = new PlotArea();
        var layout = new Layout();
        plotArea.Append(layout);

        var barChartElement = new DocumentFormat.OpenXml.Drawing.Charts.BarChart();
        barChartElement.Append(new BarDirection { Val = barChart.Orientation == BarChartOrientation.Horizontal 
            ? BarDirectionValues.Bar 
            : BarDirectionValues.Column });
        barChartElement.Append(new BarGrouping { Val = BarGroupingValues.Clustered });
        barChartElement.Append(new VaryColors { Val = false });

        if (barChart.ValuesRange != null)
        {
            var barChartSeries = new BarChartSeries();
            barChartSeries.Append(new Index { Val = 0 });
            barChartSeries.Append(new Order { Val = 0 });

            var seriesText = new SeriesText();
            var stringValue = new NumericValue { Text = chart.SingleSeriesName ?? "Series 1" };
            seriesText.Append(stringValue);
            barChartSeries.Append(seriesText);

            if (barChart.CategoriesRange != null)
            {
                var categoryAxisData = new CategoryAxisData();
                var catRef = new StringReference();
                catRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(barChart.CategoriesRange.Value, sheetName) });
                categoryAxisData.Append(catRef);
                barChartSeries.Append(categoryAxisData);
            }

            var values = new Values();
            var numRef = new NumberReference();
            numRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(barChart.ValuesRange.Value, sheetName) });
            values.Append(numRef);
            barChartSeries.Append(values);

            AddSeriesColor(barChartSeries, chart.SingleSeriesColor, includeFill: true, includeLine: false);

            barChartElement.Append(barChartSeries);
        }

        if (chart.Series.Count != 0)
            for (var i = 0; i < chart.Series.Count; i++)
            {
                var series = chart.Series[i];
                var barChartSeries = new BarChartSeries();
                barChartSeries.Append(new Index { Val = (uint)i });
                barChartSeries.Append(new Order { Val = (uint)i });

                var seriesText = new SeriesText();
                var stringValue = new NumericValue { Text = series.Name };
                seriesText.Append(stringValue);
                barChartSeries.Append(seriesText);

                var values = new Values();
                var numRef = new NumberReference();
                numRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(series.DataRange, sheetName) });
                values.Append(numRef);
                barChartSeries.Append(values);

                AddSeriesColor(barChartSeries, series.Color, includeFill: true, includeLine: false);

                barChartElement.Append(barChartSeries);
            }

        barChartElement.Append(new DataLabels(new ShowLegendKey { Val = false }, 
            new ShowValue { Val = chart.ShowDataLabels }, 
            new ShowCategoryName { Val = false }, 
            new ShowSeriesName { Val = false }, 
            new ShowPercent { Val = false }, 
            new ShowBubbleSize { Val = false }));

        barChartElement.Append(new AxisId { Val = 1 });
        barChartElement.Append(new AxisId { Val = 2 });

        plotArea.Append(barChartElement);

        var categoryAxis = new CategoryAxis();
        categoryAxis.Append(new AxisId { Val = 1 });
        categoryAxis.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        categoryAxis.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        categoryAxis.Append(new CrossingAxis { Val = 2 });
        categoryAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
            categoryAxis.Append(CreateAxisTitle(chart.CategoryAxisTitle));
        categoryAxis.Append(new AutoLabeled { Val = true });
        categoryAxis.Append(new LabelAlignment { Val = LabelAlignmentValues.Center });
        categoryAxis.Append(new LabelOffset { Val = 100 });
        plotArea.Append(categoryAxis);

        var valueAxis = new ValueAxis();
        valueAxis.Append(new AxisId { Val = 2 });
        valueAxis.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        valueAxis.Append(new Delete { Val = false });
        valueAxis.Append(new AxisPosition { Val = AxisPositionValues.Left });
        if (chart.ShowMajorGridlines)
            valueAxis.Append(new MajorGridlines());
        if (!chart.ShowYAxisLabels)
            valueAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.None });
        else
        {
            valueAxis.Append(new NumberingFormat 
            { 
                FormatCode = "General", 
                SourceLinked = true 
            });
            valueAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        }
        valueAxis.Append(new CrossingAxis { Val = 1 });
        valueAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
            valueAxis.Append(CreateAxisTitle(chart.ValueAxisTitle));
        valueAxis.Append(new CrossBetween { Val = CrossBetweenValues.Between });
        plotArea.Append(valueAxis);

        chartElement.Append(plotArea);

        if (chart.LegendPosition != ChartLegendPosition.None)
        {
            var legend = new Legend();
            legend.Append(new LegendPosition { Val = GetLegendPositionValue(chart.LegendPosition) });
            legend.Append(new Layout());
            legend.Append(new Overlay { Val = false });
            chartElement.Append(legend);
        }

        chartElement.Append(new PlotVisibleOnly { Val = true });
        chartElement.Append(new DisplayBlanksAs { Val = DisplayBlanksAsValues.Gap });
        chartElement.Append(new ShowDataLabelsOverMaximum { Val = false });

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new DefaultRunProperties());
            paragraph.Append(paragraphProperties);
            
            var run = new Run();
            run.Append(new RunProperties { Language = "en-US" });
            run.Append(new Text { Text = chart.Title });
            paragraph.Append(run);
            
            richText.Append(paragraph);
            chartText.Append(richText);
            title.Append(chartText);
            title.Append(new Layout());
            title.Append(new Overlay { Val = false });
            
            chartElement.InsertAt(title, 0);
        }

        chartSpace.Append(new PrintSettings(
            new HeaderFooter(),
            new PageMargins { Left = 0.7, Right = 0.7, Top = 0.75, Bottom = 0.75, Header = 0.3, Footer = 0.3 }
        ));

        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();
    }

    private static void GenerateLineChartPart(ChartPart chartPart, LineChart lineChart, string sheetName, Chart chart)
    {
        var chartSpace = new ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var editingLanguage = new EditingLanguage { Val = "en-US" };
        chartSpace.Append(editingLanguage);

        var chartElement = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        
        var plotArea = new PlotArea();
        var layout = new Layout();
        plotArea.Append(layout);

        var lineChartElement = new DocumentFormat.OpenXml.Drawing.Charts.LineChart();
        lineChartElement.Append(new Grouping { Val = GroupingValues.Standard });
        lineChartElement.Append(new VaryColors { Val = false });

        if (lineChart.ValuesRange != null)
        {
            var lineChartSeries = new LineChartSeries();
            lineChartSeries.Append(new Index { Val = 0 });
            lineChartSeries.Append(new Order { Val = 0 });

            var seriesText = new SeriesText();
            var stringValue = new NumericValue { Text = chart.SingleSeriesName ?? "Series 1" };
            seriesText.Append(stringValue);
            lineChartSeries.Append(seriesText);

            var marker = new Marker();
            marker.Append(new Symbol { Val = ConvertLineChartMarkerStyle(lineChart.MarkerStyle) });
            if (lineChart.MarkerStyle != LineChartMarkerStyle.None)
                marker.Append(new Size { Val = (byte)5 });
            lineChartSeries.Append(marker);

            if (lineChart.CategoriesRange != null)
            {
                var categoryAxisData = new CategoryAxisData();
                var catRef = new StringReference();
                catRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(lineChart.CategoriesRange.Value, sheetName) });
                categoryAxisData.Append(catRef);
                lineChartSeries.Append(categoryAxisData);
            }

            var values = new Values();
            var numRef = new NumberReference();
            numRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(lineChart.ValuesRange.Value, sheetName) });
            values.Append(numRef);
            lineChartSeries.Append(values);

            AddSeriesColor(lineChartSeries, chart.SingleSeriesColor, includeFill: false, includeLine: true);

            lineChartElement.Append(lineChartSeries);
        }

        if (chart.Series.Count != 0)
            for (var i = 0; i < chart.Series.Count; i++)
            {
                var series = chart.Series[i];
                var lineChartSeries = new LineChartSeries();
                lineChartSeries.Append(new Index { Val = (uint)i });
                lineChartSeries.Append(new Order { Val = (uint)i });

                var seriesText = new SeriesText();
                var stringValue = new NumericValue { Text = series.Name };
                seriesText.Append(stringValue);
                lineChartSeries.Append(seriesText);

                var marker = new Marker();
                marker.Append(new Symbol { Val = ConvertLineChartMarkerStyle(lineChart.MarkerStyle) });
                if (lineChart.MarkerStyle != LineChartMarkerStyle.None)
                    marker.Append(new Size { Val = (byte)5 });
                lineChartSeries.Append(marker);

                var values = new Values();
                var numRef = new NumberReference();
                numRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(series.DataRange, sheetName) });
                values.Append(numRef);
                lineChartSeries.Append(values);

                AddSeriesColor(lineChartSeries, series.Color, includeFill: false, includeLine: true);

                lineChartElement.Append(lineChartSeries);
            }

        lineChartElement.Append(new DataLabels(
            new ShowLegendKey { Val = false }, 
            new ShowValue { Val = chart.ShowDataLabels }, 
            new ShowCategoryName { Val = false }, 
            new ShowSeriesName { Val = false }, 
            new ShowPercent { Val = false }, 
            new ShowBubbleSize { Val = false }));

        lineChartElement.Append(new AxisId { Val = 1 });
        lineChartElement.Append(new AxisId { Val = 2 });

        plotArea.Append(lineChartElement);

        var categoryAxis = new CategoryAxis();
        categoryAxis.Append(new AxisId { Val = 1 });
        categoryAxis.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        categoryAxis.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        categoryAxis.Append(new CrossingAxis { Val = 2 });
        categoryAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
            categoryAxis.Append(CreateAxisTitle(chart.CategoryAxisTitle));
        plotArea.Append(categoryAxis);

        var valueAxis = new ValueAxis();
        valueAxis.Append(new AxisId { Val = 2 });
        valueAxis.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        valueAxis.Append(new Delete { Val = false });
        valueAxis.Append(new AxisPosition { Val = AxisPositionValues.Left });
        if (chart.ShowMajorGridlines)
            valueAxis.Append(new MajorGridlines());
        if (!chart.ShowYAxisLabels)
            valueAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.None });
        else
        {
            valueAxis.Append(new NumberingFormat { FormatCode = "General", SourceLinked = true });
            valueAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        }
        valueAxis.Append(new CrossingAxis { Val = 1 });
        valueAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
            valueAxis.Append(CreateAxisTitle(chart.ValueAxisTitle));
        plotArea.Append(valueAxis);

        chartElement.Append(plotArea);

        if (chart.LegendPosition != ChartLegendPosition.None)
        {
            var legend = new Legend();
            legend.Append(new LegendPosition { Val = GetLegendPositionValue(chart.LegendPosition) });
            legend.Append(new Layout());
            legend.Append(new Overlay { Val = false });
            chartElement.Append(legend);
        }

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties { Language = "en-US" });
            run.Append(new Text { Text = chart.Title });
            paragraph.Append(run);
            
            richText.Append(paragraph);
            chartText.Append(richText);
            title.Append(chartText);
            title.Append(new Layout());
            title.Append(new Overlay { Val = false });
            
            chartElement.InsertAt(title, 0);
        }

        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();
    }

    private static void GenerateAreaChartPart(ChartPart chartPart, AreaChart areaChart, string sheetName, Chart chart)
    {
        var chartSpace = new ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var editingLanguage = new EditingLanguage { Val = "en-US" };
        chartSpace.Append(editingLanguage);

        var chartElement = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        
        var plotArea = new PlotArea();
        var layout = new Layout();
        plotArea.Append(layout);

        var areaChartElement = new DocumentFormat.OpenXml.Drawing.Charts.AreaChart();
        var grouping = new Grouping { Val = areaChart.Stacked ? GroupingValues.Stacked : GroupingValues.Standard };
        areaChartElement.Append(grouping);
        areaChartElement.Append(new VaryColors { Val = false });

        if (areaChart.ValuesRange != null)
        {
            var areaChartSeries = new AreaChartSeries();
            areaChartSeries.Append(new Index { Val = 0 });
            areaChartSeries.Append(new Order { Val = 0 });

            var seriesText = new SeriesText();
            var stringValue = new NumericValue { Text = chart.SingleSeriesName ?? "Series 1" };
            seriesText.Append(stringValue);
            areaChartSeries.Append(seriesText);

            if (areaChart.CategoriesRange != null)
            {
                var categoryAxisData = new CategoryAxisData();
                var catRef = new StringReference();
                catRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(areaChart.CategoriesRange.Value, sheetName) });
                categoryAxisData.Append(catRef);
                areaChartSeries.Append(categoryAxisData);
            }

            var values = new Values();
            var numRef = new NumberReference();
            numRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(areaChart.ValuesRange.Value, sheetName) });
            values.Append(numRef);
            areaChartSeries.Append(values);

            AddSeriesColor(areaChartSeries, chart.SingleSeriesColor, includeFill: true, includeLine: false);

            areaChartElement.Append(areaChartSeries);
        }

        if (chart.Series.Count != 0)
            for (var i = 0; i < chart.Series.Count; i++)
            {
                var series = chart.Series[i];
                var areaChartSeries = new AreaChartSeries();
                areaChartSeries.Append(new Index { Val = (uint)(i + (areaChart.CategoriesRange != null ? 1 : 0)) });
                areaChartSeries.Append(new Order { Val = (uint)(i + (areaChart.CategoriesRange != null ? 1 : 0)) });

                var seriesText = new SeriesText();
                var stringValue = new NumericValue { Text = series.Name };
                seriesText.Append(stringValue);
                areaChartSeries.Append(seriesText);

                var values = new Values();
                var numRef = new NumberReference();
                numRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(series.DataRange, sheetName) });
                values.Append(numRef);
                areaChartSeries.Append(values);

                AddSeriesColor(areaChartSeries, series.Color, includeFill: true, includeLine: false);

                areaChartElement.Append(areaChartSeries);
            }

        areaChartElement.Append(new AxisId { Val = 1 });
        areaChartElement.Append(new AxisId { Val = 2 });

        plotArea.Append(areaChartElement);

        var categoryAxis = new CategoryAxis();
        categoryAxis.Append(new AxisId { Val = 1 });
        categoryAxis.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        categoryAxis.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        categoryAxis.Append(new CrossingAxis { Val = 2 });
        categoryAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        plotArea.Append(categoryAxis);

        var valueAxis = new ValueAxis();
        valueAxis.Append(new AxisId { Val = 2 });
        valueAxis.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        valueAxis.Append(new Delete { Val = false });
        valueAxis.Append(new AxisPosition { Val = AxisPositionValues.Left });
        if (chart.ShowMajorGridlines)
            valueAxis.Append(new MajorGridlines());
        if (!chart.ShowYAxisLabels)
            valueAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.None });
        else
        {
            valueAxis.Append(new NumberingFormat { FormatCode = "General", SourceLinked = true });
            valueAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        }
        valueAxis.Append(new CrossingAxis { Val = 1 });
        valueAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
            valueAxis.Append(CreateAxisTitle(chart.ValueAxisTitle));
        plotArea.Append(valueAxis);

        chartElement.Append(plotArea);

        if (chart.LegendPosition != ChartLegendPosition.None)
        {
            var legend = new Legend();
            legend.Append(new LegendPosition { Val = GetLegendPositionValue(chart.LegendPosition) });
            legend.Append(new Layout());
            legend.Append(new Overlay { Val = false });
            chartElement.Append(legend);
        }

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties { Language = "en-US" });
            run.Append(new Text { Text = chart.Title });
            paragraph.Append(run);
            
            richText.Append(paragraph);
            chartText.Append(richText);
            title.Append(chartText);
            title.Append(new Layout());
            title.Append(new Overlay { Val = false });
            
            chartElement.InsertAt(title, 0);
        }

        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();
    }

    private static void GeneratePieChartPart(ChartPart chartPart, PieChart pieChart, string sheetName, Chart chart)
    {
        var chartSpace = new ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var editingLanguage = new EditingLanguage { Val = "en-US" };
        chartSpace.Append(editingLanguage);

        var chartElement = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        
        var plotArea = new PlotArea();
        var layout = new Layout();
        plotArea.Append(layout);

        var pieChartElement = new DocumentFormat.OpenXml.Drawing.Charts.PieChart();
        pieChartElement.Append(new VaryColors { Val = true });

        if (pieChart is { CategoriesRange: not null, ValuesRange: not null })
        {
            var pieChartSeries = new PieChartSeries();
            pieChartSeries.Append(new Index { Val = 0 });
            pieChartSeries.Append(new Order { Val = 0 });

            var seriesText = new SeriesText();
            var stringValue = new NumericValue { Text = chart.SingleSeriesName ?? "Series 1" };
            seriesText.Append(stringValue);
            pieChartSeries.Append(seriesText);

            var categoryAxisData = new CategoryAxisData();
            var catRef = new StringReference();
            catRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(pieChart.CategoriesRange.Value, sheetName) });
            categoryAxisData.Append(catRef);
            pieChartSeries.Append(categoryAxisData);

            var values = new Values();
            var numRef = new NumberReference();
            numRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(pieChart.ValuesRange.Value, sheetName) });
            values.Append(numRef);
            pieChartSeries.Append(values);

            AddSeriesColor(pieChartSeries, chart.SingleSeriesColor, includeFill: true, includeLine: false);

            pieChartElement.Append(pieChartSeries);
        }

        pieChartElement.Append(new DataLabels(
            new ShowLegendKey { Val = false }, 
            new ShowValue { Val = chart.ShowDataLabels }, 
            new ShowCategoryName { Val = false }, 
            new ShowSeriesName { Val = false }, 
            new ShowPercent { Val = false }, 
            new ShowBubbleSize { Val = false }));

        plotArea.Append(pieChartElement);
        chartElement.Append(plotArea);

        if (chart.LegendPosition != ChartLegendPosition.None)
        {
            var legend = new Legend();
            legend.Append(new LegendPosition { Val = GetLegendPositionValue(chart.LegendPosition) });
            legend.Append(new Layout());
            legend.Append(new Overlay { Val = false });
            chartElement.Append(legend);
        }

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties { Language = "en-US" });
            run.Append(new Text { Text = chart.Title });
            paragraph.Append(run);
            
            richText.Append(paragraph);
            chartText.Append(richText);
            title.Append(chartText);
            title.Append(new Layout());
            title.Append(new Overlay { Val = false });
            
            chartElement.InsertAt(title, 0);
        }

        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();
    }

    private static void GenerateScatterChartPart(ChartPart chartPart, ScatterChart scatterChart, string sheetName, Chart chart)
    {
        var chartSpace = new ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var editingLanguage = new EditingLanguage { Val = "en-US" };
        chartSpace.Append(editingLanguage);

        var chartElement = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        
        var plotArea = new PlotArea();
        var layout = new Layout();
        plotArea.Append(layout);

        var scatterChartElement = new DocumentFormat.OpenXml.Drawing.Charts.ScatterChart();
        scatterChartElement.Append(new ScatterStyle { Val = ScatterStyleValues.LineMarker });
        scatterChartElement.Append(new VaryColors { Val = false });

        if (scatterChart is { XRange: not null, YRange: not null })
        {
            var scatterChartSeries = new ScatterChartSeries();
            scatterChartSeries.Append(new Index { Val = 0 });
            scatterChartSeries.Append(new Order { Val = 0 });

            var seriesText = new SeriesText();
            var stringValue = new NumericValue { Text = chart.SingleSeriesName ?? "Series 1" };
            seriesText.Append(stringValue);
            scatterChartSeries.Append(seriesText);

            var xValues = new XValues();
            var xNumRef = new NumberReference();
            xNumRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(scatterChart.XRange.Value, sheetName) });
            xValues.Append(xNumRef);
            scatterChartSeries.Append(xValues);

            var yValues = new YValues();
            var yNumRef = new NumberReference();
            yNumRef.Append(new Formula { Text = ChartDataRange.ToRangeReference(scatterChart.YRange.Value, sheetName) });
            yValues.Append(yNumRef);
            scatterChartSeries.Append(yValues);

            AddSeriesColor(scatterChartSeries, chart.SingleSeriesColor, includeFill: false, includeLine: true);

            scatterChartElement.Append(scatterChartSeries);
        }

        scatterChartElement.Append(new DataLabels(
            new ShowLegendKey { Val = false }, 
            new ShowValue { Val = chart.ShowDataLabels }, 
            new ShowCategoryName { Val = false }, 
            new ShowSeriesName { Val = false }, 
            new ShowPercent { Val = false }, 
            new ShowBubbleSize { Val = false }));

        scatterChartElement.Append(new AxisId { Val = 1 });
        scatterChartElement.Append(new AxisId { Val = 2 });

        plotArea.Append(scatterChartElement);

        var valueAxis1 = new ValueAxis();
        valueAxis1.Append(new AxisId { Val = 1 });
        valueAxis1.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        valueAxis1.Append(new Delete { Val = false });
        valueAxis1.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        valueAxis1.Append(new CrossingAxis { Val = 2 });
        valueAxis1.Append(new Crosses { Val = CrossesValues.AutoZero });
        if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
            valueAxis1.Append(CreateAxisTitle(chart.CategoryAxisTitle));
        plotArea.Append(valueAxis1);

        var valueAxis2 = new ValueAxis();
        valueAxis2.Append(new AxisId { Val = 2 });
        valueAxis2.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
        valueAxis2.Append(new Delete { Val = false });
        valueAxis2.Append(new AxisPosition { Val = AxisPositionValues.Left });
        if (chart.ShowMajorGridlines)
            valueAxis2.Append(new MajorGridlines());
        if (!chart.ShowYAxisLabels)
            valueAxis2.Append(new TickLabelPosition { Val = TickLabelPositionValues.None });
        else
        {
            valueAxis2.Append(new NumberingFormat { FormatCode = "General", SourceLinked = true });
            valueAxis2.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        }
        valueAxis2.Append(new CrossingAxis { Val = 1 });
        valueAxis2.Append(new Crosses { Val = CrossesValues.AutoZero });
        if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
            valueAxis2.Append(CreateAxisTitle(chart.ValueAxisTitle));
        plotArea.Append(valueAxis2);

        chartElement.Append(plotArea);

        if (chart.LegendPosition != ChartLegendPosition.None)
        {
            var legend = new Legend();
            legend.Append(new LegendPosition { Val = GetLegendPositionValue(chart.LegendPosition) });
            legend.Append(new Layout());
            legend.Append(new Overlay { Val = false });
            chartElement.Append(legend);
        }

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties { Language = "en-US" });
            run.Append(new Text { Text = chart.Title });
            paragraph.Append(run);
            
            richText.Append(paragraph);
            chartText.Append(richText);
            title.Append(chartText);
            title.Append(new Layout());
            title.Append(new Overlay { Val = false });
            
            chartElement.InsertAt(title, 0);
        }

        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();
    }

    private static TwoCellAnchor CreateTwoCellAnchor(Chart chart, string chartRelId, uint chartIndex)
    {
        if (chart.Position == null)
            throw new InvalidOperationException("Chart must have a position");

        var twoCellAnchor = new TwoCellAnchor { EditAs = EditAsValues.OneCell };

        var fromMarker = new FromMarker();
        fromMarker.Append(new ColumnId { Text = chart.Position.FromColumn.ToString() });
        fromMarker.Append(new ColumnOffset { Text = "0" });
        fromMarker.Append(new RowId { Text = chart.Position.FromRow.ToString() });
        fromMarker.Append(new RowOffset { Text = "0" });
        twoCellAnchor.Append(fromMarker);

        var toMarker = new ToMarker();
        toMarker.Append(new ColumnId { Text = chart.Position.ToColumn.ToString() });
        toMarker.Append(new ColumnOffset { Text = "0" });
        toMarker.Append(new RowId { Text = chart.Position.ToRow.ToString() });
        toMarker.Append(new RowOffset { Text = "0" });
        twoCellAnchor.Append(toMarker);

        var graphicFrame = new GraphicFrame();
        graphicFrame.Append(new NonVisualGraphicFrameProperties(
            new NonVisualDrawingProperties { Id = chartIndex, Name = $"Chart {chartIndex}" },
            new NonVisualGraphicFrameDrawingProperties()
        ));

        var transform = new Transform();
        transform.Append(new Offset { X = 0, Y = 0 });
        transform.Append(new Extents { Cx = chart.Size.WidthEmus, Cy = chart.Size.HeightEmus });
        graphicFrame.Append(transform);

        var graphic = new Graphic();
        var graphicData = new GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
        var chartReference = new ChartReference { Id = chartRelId };
        graphicData.Append(chartReference);
        graphic.Append(graphicData);
        graphicFrame.Append(graphic);

        twoCellAnchor.Append(graphicFrame);
        twoCellAnchor.Append(new ClientData());

        return twoCellAnchor;
    }

    private static OneCellAnchor CreateImageTwoCellAnchor(WorksheetImage image, string imageRelId, uint imageIndex)
    {
        var oneCellAnchor = new OneCellAnchor();

        var fromMarker = new FromMarker();
        fromMarker.Append(new ColumnId { Text = image.Position.X.ToString() });
        fromMarker.Append(new ColumnOffset { Text = "0" });
        fromMarker.Append(new RowId { Text = image.Position.Y.ToString() });
        fromMarker.Append(new RowOffset { Text = "0" });
        oneCellAnchor.Append(fromMarker);

        const int emusPerPixel = 9525;
        var extent = new Extent
        {
            Cx = image.WidthInPixels * emusPerPixel,
            Cy = image.HeightInPixels * emusPerPixel
        };
        oneCellAnchor.Append(extent);

        var picture = new Picture();
        picture.Append(new NonVisualPictureProperties(
            new NonVisualDrawingProperties 
            { 
                Id = imageIndex, 
                Name = $"Image {imageIndex}" 
            },
            new NonVisualPictureDrawingProperties(
                new PictureLocks { NoChangeAspect = true }
            )
        ));

        var blipFill = new BlipFill();
        var blip = new Blip { Embed = imageRelId };
        blip.Append(new BlipExtensionList());
        blipFill.Append(blip);
        blipFill.Append(new Stretch(new FillRectangle()));
        picture.Append(blipFill);

        var shapeProperties = new ShapeProperties();
        var transform2D = new Transform2D();
        transform2D.Append(new Offset { X = 0, Y = 0 });
        transform2D.Append(new Extents { Cx = image.WidthInPixels * emusPerPixel, Cy = image.HeightInPixels * emusPerPixel });
        shapeProperties.Append(transform2D);
        
        var presetGeometry = new PresetGeometry { Preset = ShapeTypeValues.Rectangle };
        presetGeometry.Append(new AdjustValueList());
        shapeProperties.Append(presetGeometry);
        
        picture.Append(shapeProperties);

        oneCellAnchor.Append(picture);
        oneCellAnchor.Append(new ClientData());

        return oneCellAnchor;
    }

    private static LegendPositionValues GetLegendPositionValue(ChartLegendPosition position) => position switch
    {
        ChartLegendPosition.Top => LegendPositionValues.Top,
        ChartLegendPosition.Bottom => LegendPositionValues.Bottom,
        ChartLegendPosition.Left => LegendPositionValues.Left,
        _ => LegendPositionValues.Right
    };

    private static Title CreateAxisTitle(string titleText)
    {
        var title = new Title();
        var chartText = new ChartText();
        var richText = new RichText();
        richText.Append(new BodyProperties());
        richText.Append(new ListStyle());
        
        var paragraph = new Paragraph();
        var run = new Run();
        run.Append(new RunProperties { Language = "en-US" });
        run.Append(new Text { Text = titleText });
        paragraph.Append(run);
        
        richText.Append(paragraph);
        chartText.Append(richText);
        title.Append(chartText);
        title.Append(new Layout());
        title.Append(new Overlay { Val = false });
        
        return title;
    }

    private static MarkerStyleValues ConvertLineChartMarkerStyle(LineChartMarkerStyle style) => style switch
    {
        LineChartMarkerStyle.None => MarkerStyleValues.None,
        LineChartMarkerStyle.Circle => MarkerStyleValues.Circle,
        LineChartMarkerStyle.Diamond => MarkerStyleValues.Diamond,
        _ => MarkerStyleValues.None
    };
}


