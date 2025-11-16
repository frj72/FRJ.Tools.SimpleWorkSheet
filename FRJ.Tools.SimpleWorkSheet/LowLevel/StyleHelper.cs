using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using Cell = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.Cell;
using Colors = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.Colors;
using Underline = DocumentFormat.OpenXml.Drawing.Underline;

namespace FRJ.Tools.SimpleWorkSheet.LowLevel;

public class StyleHelper
{
    private readonly List<Font> _fonts =  [new()];
    private readonly List<Fill> _fills = [new(new PatternFill { PatternType = PatternValues.None }), new(new PatternFill { PatternType = PatternValues.Gray125 }) ];
    private readonly Dictionary<string, uint> _fillIndexDictionary = new();
    private readonly List<Border> _borders = [new()];
    private readonly NumberingFormatsProvider _numberingFormatsProvider = new ();

    private readonly List<CellFormat> _cellFormats =
    [
        new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 0,
            ApplyFont = false,
            ApplyFill = false,
            ApplyBorder = false
        }];
    private readonly Dictionary<StyleDefinition, uint> _styleIndexDictionary = new();

    

    public void CollectStyles(WorkBook workBook)
    {
        foreach (var workSheet in workBook.Sheets)
        foreach (var cellEntry in workSheet.Cells.Cells)
        {
            var cell = cellEntry.Value;

            var styleDef = new StyleDefinition
            {
                FillColor = cell.Color,
                Font = cell.Font ?? WorkSheetDefaults.Font,
                Borders = cell.Borders ?? WorkSheetDefaults.CellBorders, 
                FormatAsDate = cell.Value.CellValueType() == CellValueBasicType.DateType,
                HorizontalAlignment = cell.Style.HorizontalAlignment,
                VerticalAlignment = cell.Style.VerticalAlignment,
                TextRotation = cell.Style.TextRotation,
                WrapText = cell.Style.WrapText
            };

            if (_styleIndexDictionary.ContainsKey(styleDef)) continue;
            var fontId = AddFont(styleDef.Font);
            var fillId = AddFill(styleDef.FillColor);
            var borderId = AddBorder(styleDef.Borders);

            var hasAlignment = styleDef.HorizontalAlignment.HasValue || 
                              styleDef.VerticalAlignment.HasValue || 
                              styleDef.TextRotation.HasValue || 
                              styleDef.WrapText.HasValue;

            var cellFormat = new CellFormat
            {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                ApplyFont = fontId != 0,
                ApplyFill = fillId != 0,
                ApplyBorder = borderId != 0,
                NumberFormatId = styleDef.FormatAsDate ? 164 : null,
                ApplyNumberFormat = styleDef.FormatAsDate ? true : null,
                ApplyAlignment = hasAlignment ? true : null
            };

            if (hasAlignment)
                cellFormat.Alignment = CreateAlignment(styleDef);
                    
            _cellFormats.Add(cellFormat);
            var styleIndex = (uint)_cellFormats.Count - 1;
            _styleIndexDictionary.Add(styleDef, styleIndex);
        }
    }




    public Stylesheet GenerateStylesheet()
    {
        var stylesheet = new Stylesheet
        {
            Fonts = new() { Count = (uint)_fonts.Count },
            Fills = new() { Count = (uint)_fills.Count },
            Borders = new() { Count = (uint)_borders.Count },
            CellFormats = new() { Count = (uint)_cellFormats.Count },
            NumberingFormats = _numberingFormatsProvider.NumberingFormats()
        };

        stylesheet.Fonts.Append(_fonts);
        stylesheet.Fills.Append(_fills);
        stylesheet.Borders.Append(_borders);
        stylesheet.CellFormats.Append(_cellFormats);
        return stylesheet;
    }

    public uint GetStyleIndex(Cell cell) =>
        _styleIndexDictionary[new()
        {
            FillColor = cell.Color,
            Font = cell.Font ?? WorkSheetDefaults.Font,
            Borders = cell.Borders ?? WorkSheetDefaults.CellBorders,
            FormatAsDate = cell.Value.IsDateTime() || cell.Value.IsDateTimeOffset(),
            HorizontalAlignment = cell.Style.HorizontalAlignment,
            VerticalAlignment = cell.Style.VerticalAlignment,
            TextRotation = cell.Style.TextRotation,
            WrapText = cell.Style.WrapText
        }];


    private uint AddFont(CellFont fontDef)
    {
        var font = new Font();

        if (fontDef.Bold)
            font.Append(new Bold());

        if (fontDef.Italic)
            font.Append(new Italic());

        if (fontDef.Underline)
            font.Append(new Underline());

        font.Append(new FontSize { Val = fontDef.Size });
        font.Append(new Color { Rgb = new() { Value = fontDef.Color.ToArgbColor() } });
        font.Append(new FontName { Val = fontDef.Name });

        _fonts.Add(font);
        return (uint)_fonts.Count - 1;
    }

    private uint AddFill(string? fillColor)
    {
        if (string.IsNullOrEmpty(fillColor))
            return 0;

        if (fillColor == Colors.White)
            return 0;

        if (_fillIndexDictionary.TryGetValue(fillColor, out var fillId))
            return fillId;
        
        var fill = new Fill();
        var patternFill = new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new() { Rgb = new(fillColor.ToArgbColor()) },
            BackgroundColor = new() { Indexed = 64 }
        };
        fill.PatternFill = patternFill;

        _fills.Add(fill);
        fillId = (uint)(_fills.Count - 1);
        _fillIndexDictionary[fillColor] = fillId;

        return fillId;
    }



    private uint AddBorder(CellBorders borderDef)
    {
        _borders.Add(new()
        {
            LeftBorder = CreateBorderStyle<LeftBorder>(borderDef.Left),
            RightBorder = CreateBorderStyle<RightBorder>(borderDef.Right),
            TopBorder = CreateBorderStyle<TopBorder>(borderDef.Top),
            BottomBorder = CreateBorderStyle<BottomBorder>(borderDef.Bottom)
        });
        return (uint)(_borders.Count - 1);
    }

    private static T CreateBorderStyle<T>(CellBorder? borderDef) where T : BorderPropertiesType, new()
    {
        var borderProp = new T();

        if (borderDef is not { Style: not CellBorderStyle.None }) return borderProp;
        borderProp.Style = MapBorderStyle(borderDef.Style);

        if (!string.IsNullOrEmpty(borderDef.Color))
            borderProp.Color = new()
            {
                Rgb = new(borderDef.Color.ToArgbColor())
            };

        return borderProp;
    }

    private static BorderStyleValues MapBorderStyle(CellBorderStyle style) =>
        style switch
        {
            CellBorderStyle.Dashed => BorderStyleValues.Dashed,
            CellBorderStyle.Dotted => BorderStyleValues.Dotted,
            CellBorderStyle.Double => BorderStyleValues.Double,
            CellBorderStyle.DashDot => BorderStyleValues.DashDot,
            CellBorderStyle.DashDotDot => BorderStyleValues.DashDotDot,
            CellBorderStyle.Hair => BorderStyleValues.Hair,
            CellBorderStyle.Medium => BorderStyleValues.Medium,
            CellBorderStyle.MediumDashDot => BorderStyleValues.MediumDashDot,
            CellBorderStyle.MediumDashDotDot => BorderStyleValues.MediumDashDotDot,
            CellBorderStyle.MediumDashed => BorderStyleValues.MediumDashed,
            CellBorderStyle.SlantDashDot => BorderStyleValues.SlantDashDot,
            CellBorderStyle.Thick => BorderStyleValues.Thick,
            CellBorderStyle.Thin => BorderStyleValues.Thin,
            _ => BorderStyleValues.None
        };

    private static Alignment CreateAlignment(StyleDefinition styleDef)
    {
        var alignment = new Alignment();

        if (styleDef.HorizontalAlignment.HasValue)
            alignment.Horizontal = MapHorizontalAlignment(styleDef.HorizontalAlignment.Value);

        if (styleDef.VerticalAlignment.HasValue)
            alignment.Vertical = MapVerticalAlignment(styleDef.VerticalAlignment.Value);

        if (styleDef.TextRotation.HasValue)
            alignment.TextRotation = (uint)styleDef.TextRotation.Value;

        if (styleDef.WrapText.HasValue)
            alignment.WrapText = styleDef.WrapText.Value;

        return alignment;
    }

    private static HorizontalAlignmentValues MapHorizontalAlignment(HorizontalAlignment alignment) =>
        alignment switch
        {
            HorizontalAlignment.Left => HorizontalAlignmentValues.Left,
            HorizontalAlignment.Center => HorizontalAlignmentValues.Center,
            HorizontalAlignment.Right => HorizontalAlignmentValues.Right,
            HorizontalAlignment.Justify => HorizontalAlignmentValues.Justify,
            HorizontalAlignment.Fill => HorizontalAlignmentValues.Fill,
            HorizontalAlignment.Distributed => HorizontalAlignmentValues.Distributed,
            _ => HorizontalAlignmentValues.Left
        };

    private static VerticalAlignmentValues MapVerticalAlignment(VerticalAlignment alignment) =>
        alignment switch
        {
            VerticalAlignment.Top => VerticalAlignmentValues.Top,
            VerticalAlignment.Middle => VerticalAlignmentValues.Center,
            VerticalAlignment.Bottom => VerticalAlignmentValues.Bottom,
            VerticalAlignment.Justify => VerticalAlignmentValues.Justify,
            VerticalAlignment.Distributed => VerticalAlignmentValues.Distributed,
            _ => VerticalAlignmentValues.Top
        };
}