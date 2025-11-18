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
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;
using Selection = DocumentFormat.OpenXml.Spreadsheet.Selection;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;
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

    private static void GenerateChartPart(ChartPart chartPart, Components.Charts.Chart chart, string sheetName)
    {
        switch (chart)
        {
            case Components.Charts.BarChart barChart:
                GenerateBarChartPart(chartPart, barChart, sheetName, chart);
                break;
            case Components.Charts.LineChart lineChart:
                GenerateLineChartPart(chartPart, lineChart, sheetName, chart);
                break;
            case Components.Charts.PieChart pieChart:
                GeneratePieChartPart(chartPart, pieChart, sheetName, chart);
                break;
            case Components.Charts.ScatterChart scatterChart:
                GenerateScatterChartPart(chartPart, scatterChart, sheetName, chart);
                break;
        }
    }

    private static void GenerateBarChartPart(ChartPart chartPart, Components.Charts.BarChart barChart, string sheetName, Components.Charts.Chart chart)
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

        if (barChart is { CategoriesRange: not null, ValuesRange: not null })
        {
            var barChartSeries = new BarChartSeries();
            barChartSeries.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0 });
            barChartSeries.Append(new Order { Val = 0 });

            var categoryAxisData = new CategoryAxisData();
            var catRef = new StringReference();
            catRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(barChart.CategoriesRange.Value, sheetName) });
            categoryAxisData.Append(catRef);
            barChartSeries.Append(categoryAxisData);

            var values = new DocumentFormat.OpenXml.Drawing.Charts.Values();
            var numRef = new NumberReference();
            numRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(barChart.ValuesRange.Value, sheetName) });
            values.Append(numRef);
            barChartSeries.Append(values);

            barChartElement.Append(barChartSeries);
        }

        barChartElement.Append(new DataLabels(new ShowLegendKey { Val = false }, 
            new ShowValue { Val = false }, 
            new ShowCategoryName { Val = false }, 
            new ShowSeriesName { Val = false }, 
            new ShowPercent { Val = false }, 
            new ShowBubbleSize { Val = false }));

        barChartElement.Append(new AxisId { Val = 1 });
        barChartElement.Append(new AxisId { Val = 2 });

        plotArea.Append(barChartElement);

        var categoryAxis = new CategoryAxis();
        categoryAxis.Append(new AxisId { Val = 1 });
        categoryAxis.Append(new Scaling(new Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
        categoryAxis.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        categoryAxis.Append(new CrossingAxis { Val = 2 });
        categoryAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        categoryAxis.Append(new AutoLabeled { Val = true });
        categoryAxis.Append(new LabelAlignment { Val = LabelAlignmentValues.Center });
        categoryAxis.Append(new LabelOffset { Val = 100 });
        plotArea.Append(categoryAxis);

        var valueAxis = new ValueAxis();
        valueAxis.Append(new AxisId { Val = 2 });
        valueAxis.Append(new Scaling(new Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
        valueAxis.Append(new AxisPosition { Val = AxisPositionValues.Left });
        valueAxis.Append(new MajorGridlines());
        valueAxis.Append(new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat 
        { 
            FormatCode = "General", 
            SourceLinked = true 
        });
        valueAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        valueAxis.Append(new CrossingAxis { Val = 1 });
        valueAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        valueAxis.Append(new CrossBetween { Val = CrossBetweenValues.Between });
        plotArea.Append(valueAxis);

        chartElement.Append(plotArea);

        var legend = new Legend();
        legend.Append(new LegendPosition { Val = LegendPositionValues.Right });
        legend.Append(new Layout());
        legend.Append(new Overlay { Val = false });
        chartElement.Append(legend);

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
            
            var run = new DocumentFormat.OpenXml.Drawing.Run();
            run.Append(new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US" });
            run.Append(new DocumentFormat.OpenXml.Drawing.Text { Text = chart.Title });
            paragraph.Append(run);
            
            richText.Append(paragraph);
            chartText.Append(richText);
            title.Append(chartText);
            title.Append(new Layout());
            title.Append(new Overlay { Val = false });
            
            chartElement.InsertAt(title, 0);
        }

        chartSpace.Append(new PrintSettings(
            new DocumentFormat.OpenXml.Drawing.Charts.HeaderFooter(),
            new DocumentFormat.OpenXml.Drawing.Charts.PageMargins { Left = 0.7, Right = 0.7, Top = 0.75, Bottom = 0.75, Header = 0.3, Footer = 0.3 }
        ));

        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();
    }

    private static void GenerateLineChartPart(ChartPart chartPart, Components.Charts.LineChart lineChart, string sheetName, Components.Charts.Chart chart)
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

        if (lineChart is { CategoriesRange: not null, ValuesRange: not null })
        {
            var lineChartSeries = new LineChartSeries();
            lineChartSeries.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0 });
            lineChartSeries.Append(new Order { Val = 0 });

            var categoryAxisData = new CategoryAxisData();
            var catRef = new StringReference();
            catRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(lineChart.CategoriesRange.Value, sheetName) });
            categoryAxisData.Append(catRef);
            lineChartSeries.Append(categoryAxisData);

            var values = new DocumentFormat.OpenXml.Drawing.Charts.Values();
            var numRef = new NumberReference();
            numRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(lineChart.ValuesRange.Value, sheetName) });
            values.Append(numRef);
            lineChartSeries.Append(values);

            lineChartElement.Append(lineChartSeries);
        }

        lineChartElement.Append(new AxisId { Val = 1 });
        lineChartElement.Append(new AxisId { Val = 2 });

        plotArea.Append(lineChartElement);

        var categoryAxis = new CategoryAxis();
        categoryAxis.Append(new AxisId { Val = 1 });
        categoryAxis.Append(new Scaling(new Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
        categoryAxis.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        categoryAxis.Append(new CrossingAxis { Val = 2 });
        categoryAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        plotArea.Append(categoryAxis);

        var valueAxis = new ValueAxis();
        valueAxis.Append(new AxisId { Val = 2 });
        valueAxis.Append(new Scaling(new Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
        valueAxis.Append(new AxisPosition { Val = AxisPositionValues.Left });
        valueAxis.Append(new MajorGridlines());
        valueAxis.Append(new CrossingAxis { Val = 1 });
        valueAxis.Append(new Crosses { Val = CrossesValues.AutoZero });
        plotArea.Append(valueAxis);

        chartElement.Append(plotArea);

        var legend = new Legend();
        legend.Append(new LegendPosition { Val = LegendPositionValues.Right });
        legend.Append(new Layout());
        legend.Append(new Overlay { Val = false });
        chartElement.Append(legend);

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var run = new DocumentFormat.OpenXml.Drawing.Run();
            run.Append(new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US" });
            run.Append(new DocumentFormat.OpenXml.Drawing.Text { Text = chart.Title });
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

    private static void GeneratePieChartPart(ChartPart chartPart, Components.Charts.PieChart pieChart, string sheetName, Components.Charts.Chart chart)
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
            pieChartSeries.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0 });
            pieChartSeries.Append(new Order { Val = 0 });

            var categoryAxisData = new CategoryAxisData();
            var catRef = new StringReference();
            catRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(pieChart.CategoriesRange.Value, sheetName) });
            categoryAxisData.Append(catRef);
            pieChartSeries.Append(categoryAxisData);

            var values = new DocumentFormat.OpenXml.Drawing.Charts.Values();
            var numRef = new NumberReference();
            numRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(pieChart.ValuesRange.Value, sheetName) });
            values.Append(numRef);
            pieChartSeries.Append(values);

            pieChartElement.Append(pieChartSeries);
        }

        plotArea.Append(pieChartElement);
        chartElement.Append(plotArea);

        var legend = new Legend();
        legend.Append(new LegendPosition { Val = LegendPositionValues.Right });
        legend.Append(new Layout());
        legend.Append(new Overlay { Val = false });
        chartElement.Append(legend);

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var run = new DocumentFormat.OpenXml.Drawing.Run();
            run.Append(new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US" });
            run.Append(new DocumentFormat.OpenXml.Drawing.Text { Text = chart.Title });
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

    private static void GenerateScatterChartPart(ChartPart chartPart, Components.Charts.ScatterChart scatterChart, string sheetName, Components.Charts.Chart chart)
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
            scatterChartSeries.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0 });
            scatterChartSeries.Append(new Order { Val = 0 });

            var xValues = new XValues();
            var xNumRef = new NumberReference();
            xNumRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(scatterChart.XRange.Value, sheetName) });
            xValues.Append(xNumRef);
            scatterChartSeries.Append(xValues);

            var yValues = new YValues();
            var yNumRef = new NumberReference();
            yNumRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = ChartDataRange.ToRangeReference(scatterChart.YRange.Value, sheetName) });
            yValues.Append(yNumRef);
            scatterChartSeries.Append(yValues);

            scatterChartElement.Append(scatterChartSeries);
        }

        scatterChartElement.Append(new AxisId { Val = 1 });
        scatterChartElement.Append(new AxisId { Val = 2 });

        plotArea.Append(scatterChartElement);

        var valueAxis1 = new ValueAxis();
        valueAxis1.Append(new AxisId { Val = 1 });
        valueAxis1.Append(new Scaling(new Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
        valueAxis1.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        valueAxis1.Append(new CrossingAxis { Val = 2 });
        valueAxis1.Append(new Crosses { Val = CrossesValues.AutoZero });
        plotArea.Append(valueAxis1);

        var valueAxis2 = new ValueAxis();
        valueAxis2.Append(new AxisId { Val = 2 });
        valueAxis2.Append(new Scaling(new Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
        valueAxis2.Append(new AxisPosition { Val = AxisPositionValues.Left });
        valueAxis2.Append(new MajorGridlines());
        valueAxis2.Append(new CrossingAxis { Val = 1 });
        valueAxis2.Append(new Crosses { Val = CrossesValues.AutoZero });
        plotArea.Append(valueAxis2);

        chartElement.Append(plotArea);

        var legend = new Legend();
        legend.Append(new LegendPosition { Val = LegendPositionValues.Right });
        legend.Append(new Layout());
        legend.Append(new Overlay { Val = false });
        chartElement.Append(legend);

        chartSpace.Append(chartElement);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            var title = new Title();
            var chartText = new ChartText();
            var richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            
            var paragraph = new Paragraph();
            var run = new DocumentFormat.OpenXml.Drawing.Run();
            run.Append(new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US" });
            run.Append(new DocumentFormat.OpenXml.Drawing.Text { Text = chart.Title });
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

    private static TwoCellAnchor CreateTwoCellAnchor(Components.Charts.Chart chart, string chartRelId, uint chartIndex)
    {
        if (chart.Position == null)
            throw new InvalidOperationException("Chart must have a position");

        var twoCellAnchor = new TwoCellAnchor { EditAs = EditAsValues.OneCell };

        var fromMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
        fromMarker.Append(new ColumnId { Text = chart.Position.FromColumn.ToString() });
        fromMarker.Append(new ColumnOffset { Text = "0" });
        fromMarker.Append(new RowId { Text = chart.Position.FromRow.ToString() });
        fromMarker.Append(new RowOffset { Text = "0" });
        twoCellAnchor.Append(fromMarker);

        var toMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
        toMarker.Append(new ColumnId { Text = chart.Position.ToColumn.ToString() });
        toMarker.Append(new ColumnOffset { Text = "0" });
        toMarker.Append(new RowId { Text = chart.Position.ToRow.ToString() });
        toMarker.Append(new RowOffset { Text = "0" });
        twoCellAnchor.Append(toMarker);

        var graphicFrame = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame();
        graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = chartIndex, Name = $"Chart {chartIndex}" },
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()
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

        var fromMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
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

        var picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
        picture.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties(
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties 
            { 
                Id = imageIndex, 
                Name = $"Image {imageIndex}" 
            },
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties(
                new PictureLocks { NoChangeAspect = true }
            )
        ));

        var blipFill = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill();
        var blip = new Blip { Embed = imageRelId };
        blip.Append(new BlipExtensionList());
        blipFill.Append(blip);
        blipFill.Append(new Stretch(new FillRectangle()));
        picture.Append(blipFill);

        var shapeProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
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
}


