using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using OneOf;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using CellValue = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.CellValue;
using CellStyle = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.CellStyle;
using CellFormula = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.CellFormula;

namespace FRJ.Tools.SimpleWorkSheet.LowLevel;

public class WorkBookReader
{
    public static WorkBook LoadFromFile(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return LoadFromStream(stream);
    }

    public static WorkBook LoadFromStream(Stream stream)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var workbookPart = spreadsheetDocument.WorkbookPart ?? throw new InvalidOperationException("Workbook part not found");

        var workbookSheets = workbookPart.Workbook.Sheets?.Elements<Sheet>() ?? [];

        var sheets = new List<WorkSheet>();
        foreach (var sheet in workbookSheets)
        {
            var sheetId = sheet.Id?.Value;
            var sheetName = sheet.Name?.Value;
            
            if (sheetId == null || sheetName == null)
                continue;
            
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            sheets.Add(ParseWorksheet(worksheetPart, sheetName, workbookPart));
        }

        return new("ImportedWorkbook", sheets);
    }

    private static WorkSheet ParseWorksheet(WorksheetPart worksheetPart, string sheetName, WorkbookPart workbookPart)
    {
        var workSheet = new WorkSheet(sheetName);
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
        
        if (sheetData == null)
            return workSheet;

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;

        foreach (var row in sheetData.Elements<Row>())
        foreach (var cell in row.Elements<Cell>())
        {
            var position = GetCellPosition(cell.CellReference?.Value);
            if (position == null)
                continue;

            var cellValue = ExtractCellValue(cell, sharedStringTable);
            var cellStyle = ExtractCellStyle(cell, stylesheet);
            var hyperlink = ExtractHyperlink(worksheetPart, cell.CellReference?.Value);

            var simpleCell = Components.SimpleCell.Cell.Create(cellValue, cellStyle.FillColor, cellStyle.Font, cellStyle.Borders, cellStyle.FormatCode);
            
            if (cellStyle.HorizontalAlignment.HasValue || cellStyle.VerticalAlignment.HasValue || 
                cellStyle.TextRotation.HasValue || cellStyle.WrapText.HasValue)
                simpleCell = simpleCell with 
                { 
                    Style = simpleCell.Style with 
                    {
                        HorizontalAlignment = cellStyle.HorizontalAlignment,
                        VerticalAlignment = cellStyle.VerticalAlignment,
                        TextRotation = cellStyle.TextRotation,
                        WrapText = cellStyle.WrapText
                    }
                };

            if (hyperlink != null)
                simpleCell = simpleCell with { Hyperlink = hyperlink };

            workSheet.Cells.Cells[position.Value] = simpleCell;
        }

        ExtractColumnWidths(worksheetPart, workSheet);
        ExtractRowHeights(worksheetPart, workSheet);
        ExtractFrozenPanes(worksheetPart, workSheet);
        ExtractMergedCells(worksheetPart, workSheet);

        return workSheet;
    }

    private static CellPosition? GetCellPosition(string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
            return null;

        var colPart = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var rowPart = new string(cellReference.SkipWhile(char.IsLetter).ToArray());

        if (!uint.TryParse(rowPart, out var row))
            return null;

        var col = colPart.Aggregate(0u, (current, c) => current * 26 + (uint)(char.ToUpper(c) - 'A' + 1));

        return new CellPosition(col - 1, row - 1);
    }

    private static CellValue ExtractCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.CellFormula != null)
            return new CellFormula(cell.CellFormula.Text);

        var cellValue = cell.CellValue?.Text;
        if (string.IsNullOrEmpty(cellValue))
            return string.Empty;

        if (cell.DataType?.Value == CellValues.SharedString && sharedStringTable != null)
            if (int.TryParse(cellValue, out var stringIndex))
                return sharedStringTable.Elements<SharedStringItem>().ElementAt(stringIndex).InnerText;

        if (cell.DataType?.Value == CellValues.Boolean)
            return (cellValue == "1").ToString();

        if (cell.DataType?.Value == CellValues.Number || cell.DataType == null)
            if (double.TryParse(cellValue, out var doubleValue))
            {
                if (cell.StyleIndex != null)
                    return doubleValue;
                
                if (double.IsInteger(doubleValue))
                    return (long)doubleValue;
                
                return doubleValue;
            }

        if (cell.DataType?.Value == CellValues.Date && DateTime.TryParse(cellValue, out var dateValue))
            return dateValue;

        return cellValue;
    }

    private static CellStyle ExtractCellStyle(Cell cell, Stylesheet? stylesheet)
    {
        if (stylesheet == null || cell.StyleIndex == null)
            return CellStyle.Create();

        var cellFormat = stylesheet.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)cell.StyleIndex.Value);
        if (cellFormat == null)
            return CellStyle.Create();

        string? fillColor = null;
        CellFont? font = null;
        CellBorders? borders = null;
        HorizontalAlignment? horizontalAlignment = null;
        VerticalAlignment? verticalAlignment = null;
        int? textRotation = null;
        bool? wrapText = null;

        if (cellFormat.FillId != null)
        {
            var fill = stylesheet.Fills?.Elements<Fill>().ElementAtOrDefault((int)cellFormat.FillId.Value);
            if (fill?.PatternFill?.ForegroundColor?.Rgb?.Value != null)
                fillColor = NormalizeColor(fill.PatternFill.ForegroundColor.Rgb.Value);
        }

        if (cellFormat.FontId != null)
        {
            var fontDef = stylesheet.Fonts?.Elements<Font>().ElementAtOrDefault((int)cellFormat.FontId.Value);
            if (fontDef != null)
            {
                var size = (int?)(fontDef.FontSize?.Val?.Value ?? 11);
                var name = fontDef.FontName?.Val?.Value ?? "Calibri";
                var color = NormalizeColor(fontDef.Color?.Rgb?.Value ?? "FF000000");
                var bold = fontDef.Bold != null;
                var italic = fontDef.Italic != null;
                var underline = fontDef.Underline != null;

                font = CellFont.Create(size, name, color, bold, italic, underline);
            }
        }

        if (cellFormat.BorderId != null)
        {
            var border = stylesheet.Borders?.Elements<Border>().ElementAtOrDefault((int)cellFormat.BorderId.Value);
            if (border != null)
            {
                var left = ParseBorder(border.LeftBorder);
                var right = ParseBorder(border.RightBorder);
                var top = ParseBorder(border.TopBorder);
                var bottom = ParseBorder(border.BottomBorder);

                if (left != null || right != null || top != null || bottom != null)
                    borders = new(left, right, top, bottom);
            }
        }

        if (cellFormat.Alignment == null)
            return CellStyle.Create(fillColor, font, borders, null, horizontalAlignment, verticalAlignment,
                textRotation, wrapText);
        var alignment = cellFormat.Alignment;
            
        if (alignment.Horizontal != null)
            horizontalAlignment = MapHorizontalAlignment(alignment.Horizontal.Value);
            
        if (alignment.Vertical != null)
            verticalAlignment = MapVerticalAlignment(alignment.Vertical.Value);
            
        if (alignment.TextRotation != null)
        {
            var rotationValue = alignment.TextRotation.Value;
            textRotation = rotationValue <= 90 ? (int)rotationValue : (int)(rotationValue - 90);
        }
            
        if (alignment.WrapText != null)
            wrapText = alignment.WrapText.Value;

        return CellStyle.Create(fillColor, font, borders, null, horizontalAlignment, verticalAlignment, textRotation, wrapText);
    }

    private static CellBorder? ParseBorder(BorderPropertiesType? borderProp)
    {
        if (borderProp?.Style?.Value == null)
            return null;

        var style = MapBorderStyle(borderProp.Style.Value);
        var color = borderProp.Color?.Rgb?.Value;

        return new(color != null ? NormalizeColor(color) : null, style);
    }

    private static string NormalizeColor(string color)
    {
        if (color.Length == 8 && color.StartsWith("FF"))
            return color[2..];
        return color;
    }

    private static CellBorderStyle MapBorderStyle(BorderStyleValues style)
    {
        if (style == BorderStyleValues.Thin) return CellBorderStyle.Thin;
        if (style == BorderStyleValues.Medium) return CellBorderStyle.Medium;
        if (style == BorderStyleValues.Thick) return CellBorderStyle.Thick;
        if (style == BorderStyleValues.Dashed) return CellBorderStyle.Dashed;
        if (style == BorderStyleValues.Dotted) return CellBorderStyle.Dotted;
        if (style == BorderStyleValues.Double) return CellBorderStyle.Double;
        if (style == BorderStyleValues.Hair) return CellBorderStyle.Hair;
        if (style == BorderStyleValues.MediumDashed) return CellBorderStyle.MediumDashed;
        if (style == BorderStyleValues.DashDot) return CellBorderStyle.DashDot;
        if (style == BorderStyleValues.MediumDashDot) return CellBorderStyle.MediumDashDot;
        if (style == BorderStyleValues.DashDotDot) return CellBorderStyle.DashDotDot;
        if (style == BorderStyleValues.MediumDashDotDot) return CellBorderStyle.MediumDashDotDot;
        return style == BorderStyleValues.SlantDashDot ? CellBorderStyle.SlantDashDot : CellBorderStyle.None;
    }

    private static HorizontalAlignment MapHorizontalAlignment(HorizontalAlignmentValues alignment)
    {
        if (alignment == HorizontalAlignmentValues.Left) return HorizontalAlignment.Left;
        if (alignment == HorizontalAlignmentValues.Center) return HorizontalAlignment.Center;
        if (alignment == HorizontalAlignmentValues.Right) return HorizontalAlignment.Right;
        if (alignment == HorizontalAlignmentValues.Justify) return HorizontalAlignment.Justify;
        if (alignment == HorizontalAlignmentValues.Fill) return HorizontalAlignment.Fill;
        return alignment == HorizontalAlignmentValues.Distributed ? HorizontalAlignment.Distributed : HorizontalAlignment.Left;
    }

    private static VerticalAlignment MapVerticalAlignment(VerticalAlignmentValues alignment)
    {
        if (alignment == VerticalAlignmentValues.Top) return VerticalAlignment.Top;
        if (alignment == VerticalAlignmentValues.Center) return VerticalAlignment.Middle;
        if (alignment == VerticalAlignmentValues.Bottom) return VerticalAlignment.Bottom;
        if (alignment == VerticalAlignmentValues.Justify) return VerticalAlignment.Justify;
        return alignment == VerticalAlignmentValues.Distributed ? VerticalAlignment.Distributed : VerticalAlignment.Top;
    }

    private static CellHyperlink? ExtractHyperlink(WorksheetPart worksheetPart, string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
            return null;

        var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();

        var hyperlink = hyperlinks?.Elements<Hyperlink>().FirstOrDefault(h => h.Reference?.Value == cellReference);
        if (hyperlink?.Id?.Value == null)
            return null;

        var relationship = worksheetPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id.Value);
        if (relationship?.Uri == null)
            return null;

        return new(relationship.Uri.ToString(), hyperlink.Tooltip?.Value);
    }

    private static void ExtractColumnWidths(WorksheetPart worksheetPart, WorkSheet workSheet)
    {
        var columns = worksheetPart.Worksheet.Elements<Columns>().FirstOrDefault();
        if (columns == null)
            return;

        foreach (var column in columns.Elements<Column>())
        {
            if (column.Min?.Value == null || column.Width?.Value == null || column.CustomWidth?.Value != true) continue;
            var colIndex = column.Min.Value - 1;
            OneOf<double, CellWidth> width = column.Width.Value;
            workSheet.SetColumnWith(colIndex, width);
        }
    }

    private static void ExtractRowHeights(WorksheetPart worksheetPart, WorkSheet workSheet)
    {
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData == null)
            return;

        foreach (var row in sheetData.Elements<Row>())
            if (row.RowIndex?.Value != null && row.Height?.Value != null && row.CustomHeight?.Value == true)
            {
                var rowIndex = row.RowIndex.Value - 1;
                OneOf<double, RowHeight> height = row.Height.Value;
                workSheet.SetRowHeight(rowIndex, height);
            }
    }

    private static void ExtractFrozenPanes(WorksheetPart worksheetPart, WorkSheet workSheet)
    {
        var sheetViews = worksheetPart.Worksheet.Elements<SheetViews>().FirstOrDefault();
        if (sheetViews == null)
            return;

        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
        var pane = sheetView?.Elements<Pane>().FirstOrDefault();

        if (pane?.State?.Value != PaneStateValues.Frozen) return;
        var row = pane.VerticalSplit?.Value ?? 0;
        var col = pane.HorizontalSplit?.Value ?? 0;

        if (row > 0 || col > 0)
            workSheet.FreezePanes((uint)row, (uint)col);
    }

    private static void ExtractMergedCells(WorksheetPart worksheetPart, WorkSheet workSheet)
    {
        var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
        if (mergeCells == null)
            return;

        foreach (var mergeCell in mergeCells.Elements<MergeCell>())
        {
            var reference = mergeCell.Reference?.Value;
            if (string.IsNullOrEmpty(reference))
                continue;

            var parts = reference.Split(':');
            if (parts.Length != 2)
                continue;

            var start = GetCellPosition(parts[0]);
            var end = GetCellPosition(parts[1]);
            if (start == null || end == null)
                continue;

            var range = CellRange.FromPositions(start.Value, end.Value);
            workSheet.ImportMergedRange(range);
        }
    }
}
