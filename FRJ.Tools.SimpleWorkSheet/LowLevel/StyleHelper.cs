using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using Cell = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.Cell;
using Colors = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.Colors;
using Underline = DocumentFormat.OpenXml.Drawing.Underline;

namespace FRJ.Tools.SimpleWorkSheet.LowLevel;

public class StyleHelper
{
    private readonly List<Font> _fonts =  [new()];
    private readonly List<Fill> _fills = [new(new PatternFill { PatternType = PatternValues.None }), new(new PatternFill() { PatternType = PatternValues.Gray125 }) ];
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
            ApplyBorder = false,
        }];
    private readonly Dictionary<StyleDefinition, uint> _styleIndexDictionary = new();

    

    public void CollectStyles(WorkBook workBook)
    {
        foreach (var workSheet in workBook.Sheets)
        {
            foreach (var cellEntry in workSheet.Cells.Cells)
            {
                var cell = cellEntry.Value;

                var styleDef = new StyleDefinition
                {
                    FillColor = cell.Color,
                    Font = cell.Font ?? WorkSheetDefaults.Font,
                    Borders = cell.Borders ?? WorkSheetDefaults.CellBorders, 
                    FormatAsDate = cell.Value.CellValueType() == CellValueBasicType.DateType
                };

                if (_styleIndexDictionary.ContainsKey(styleDef)) continue;
                var fontId = AddFont(styleDef.Font);
                var fillId = AddFill(styleDef.FillColor);
                var borderId = AddBorder(styleDef.Borders);

                var cellFormat = new CellFormat
                {
                    FontId = fontId,
                    FillId = fillId,
                    BorderId = borderId,
                    ApplyFont = fontId != 0,
                    ApplyFill = fillId != 0,
                    ApplyBorder = borderId != 0,
                    NumberFormatId = styleDef.FormatAsDate ? 164 : null,
                    ApplyNumberFormat = styleDef.FormatAsDate ? true : null
                };
                    
                _cellFormats.Add(cellFormat);
                var styleIndex = (uint)_cellFormats.Count - 1;
                _styleIndexDictionary.Add(styleDef, styleIndex);
            }
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
            FormatAsDate = cell.Value.IsDateTime() || cell.Value.IsDateTimeOffset()
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
        font.Append(new Color { Rgb = new() { Value = fontDef.Color } });
        font.Append(new FontName { Val = fontDef.Name });

        _fonts.Add(font);
        return (uint)_fonts.Count - 1;
    }

    private uint AddFill(string? fillColor)
    {
        if (string.IsNullOrEmpty(fillColor))
            return 0;

        var fillKey = fillColor;

        if (fillKey == Colors.White)
            return 0;

        if (_fillIndexDictionary.TryGetValue(fillKey, out var fillId))
            return fillId;
        
        var fill = new Fill();
        var patternFill = new PatternFill() { PatternType = PatternValues.Solid };
        patternFill.ForegroundColor = new() { Rgb = new(fillKey) };
        patternFill.BackgroundColor = new() { Indexed = 64 };
        fill.PatternFill = patternFill;

        _fills.Add(fill);
        fillId = (uint)(_fills.Count - 1);
        _fillIndexDictionary[fillKey] = fillId;

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

    private T CreateBorderStyle<T>(CellBorder? borderDef) where T : BorderPropertiesType, new()
    {
        var borderProp = new T();

        if (borderDef is not { Style: not CellBorderStyle.None }) return borderProp;
        borderProp.Style = MapBorderStyle(borderDef.Style);

        if (!string.IsNullOrEmpty(borderDef.Color))
        {
            borderProp.Color = new()
            {
                Rgb = new(borderDef.Color)
            };
        }

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
            CellBorderStyle.None => BorderStyleValues.None,
            _ => BorderStyleValues.None
        };
}

public class NumberingFormatsProvider
{
    private readonly List<NumberingFormat> _internalNumberingFormats =
    [ 
        new() { NumberFormatId = DateFormat.IsoDateTime.ToExcelFormatId(), FormatCode = DateFormat.IsoDateTime.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.IsoDate.ToExcelFormatId(), FormatCode = DateFormat.IsoDate.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.DateTime.ToExcelFormatId(), FormatCode = DateFormat.DateTime.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.DateOnly.ToExcelFormatId(), FormatCode = DateFormat.DateOnly.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.TimeOnly.ToExcelFormatId(), FormatCode = DateFormat.TimeOnly.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Integer.ToExcelFormatId(), FormatCode = NumberFormat.Integer.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Float2.ToExcelFormatId(), FormatCode = NumberFormat.Float2.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Float3.ToExcelFormatId(), FormatCode = NumberFormat.Float3.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Float4.ToExcelFormatId(), FormatCode = NumberFormat.Float4.ToExcelFormatString() }
    ];
    
    public NumberingFormats NumberingFormats()
    {
        NumberingFormats numberingFormats = new() { Count = (uint)_internalNumberingFormats.Count };
        numberingFormats.Append(_internalNumberingFormats);
        return numberingFormats;
    }

    public uint GetOrCreateNumberFormatId(string formatCode)
    {
        var existing = _internalNumberingFormats.FirstOrDefault(nf => nf.FormatCode == formatCode);
        if (existing != null && existing.NumberFormatId != null)
            return existing.NumberFormatId;

        var newId = (uint)_internalNumberingFormats.Count + 164;
        _internalNumberingFormats.Add(new() { NumberFormatId = newId, FormatCode = formatCode });
        return newId;
    }

    public uint GetOrCreateNumberFormatId(Cell cell)
    {
        if (cell.FormatCode is not null)
            return GetOrCreateNumberFormatId(cell.FormatCode);
        
        if (cell.Value.IsDateTime() || cell.Value.IsDateTimeOffset())
            return DateFormat.IsoDateTime.ToExcelFormatId();

        if (cell.Value.IsDecimal())
            return NumberFormat.Float2.ToExcelFormatId();

        if (cell.Value.IsLong())
            return NumberFormat.Integer.ToExcelFormatId();

        return 0;
    }
}
