using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class GenericTableBuilder
{
    private readonly GenericTable _table;
    private string _sheetName = "Sheet1";
    private bool _preserveOriginalValue = true;
    private bool _trimWhitespace = true;
    private Action<CellStyleBuilder>? _headerStyleAction;
    private Dictionary<string, Func<CellValue, CellValue>>? _columnParsers;
    private bool _autoFitColumns;
    private double? _autoFitCalibration;
    private double _autoFitBaseLine;
    private List<string>? _columnOrder;
    private HashSet<string>? _excludeColumns;
    private HashSet<string>? _includeColumns;
    private DateFormat? _dateFormat;
    private Dictionary<string, NumberFormat>? _numberFormats;
    private Dictionary<string, (Func<CellValue, bool> condition, Action<CellStyleBuilder> style)>? _conditionalStyles;

    private GenericTableBuilder(GenericTable table)
    {
        ArgumentNullException.ThrowIfNull(table);
        _table = table;
    }

    public static GenericTableBuilder FromGenericTable(GenericTable table) => new(table);

    public GenericTableBuilder WithSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet name cannot be empty", nameof(name));
        
        _sheetName = name;
        return this;
    }

    public GenericTableBuilder WithPreserveOriginalValue(bool preserve)
    {
        _preserveOriginalValue = preserve;
        return this;
    }

    public GenericTableBuilder WithTrimWhitespace(bool trim)
    {
        _trimWhitespace = trim;
        return this;
    }

    public GenericTableBuilder WithHeaderStyle(Action<CellStyleBuilder> styleAction)
    {
        _headerStyleAction = styleAction ?? throw new ArgumentNullException(nameof(styleAction));
        return this;
    }

    public GenericTableBuilder WithColumnParser(string columnName, Func<CellValue, CellValue> parser)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty", nameof(columnName));

        _columnParsers ??= new();
        _columnParsers[columnName] = parser ?? throw new ArgumentNullException(nameof(parser));
        return this;
    }

    public GenericTableBuilder AutoFitAllColumns()
    {
        _autoFitColumns = true;
        _autoFitCalibration = null;
        return this;
    }
    
    public GenericTableBuilder AutoFitAllColumns(double calibration)
    {
        _autoFitColumns = true;
        _autoFitCalibration = calibration;
        return this;
    }

    public GenericTableBuilder AutoFitAllColumns(double calibration, double baseLine)
    {
        _autoFitColumns = true;
        _autoFitCalibration = calibration;
        _autoFitBaseLine = baseLine;
        return this;
    }

    public GenericTableBuilder WithColumnOrder(params string[] columnNames)
    {
        if (columnNames == null || columnNames.Length == 0)
            throw new ArgumentException("Column order must contain at least one column", nameof(columnNames));

        _columnOrder = [..columnNames];
        return this;
    }

    public GenericTableBuilder WithExcludeColumns(params string[] columnNames)
    {
        if (columnNames == null || columnNames.Length == 0)
            throw new ArgumentException("Exclude columns must contain at least one column", nameof(columnNames));

        _excludeColumns = [..columnNames];
        return this;
    }

    public GenericTableBuilder WithIncludeColumns(params string[] columnNames)
    {
        if (columnNames == null || columnNames.Length == 0)
            throw new ArgumentException("Include columns must contain at least one column", nameof(columnNames));

        _includeColumns = [..columnNames];
        return this;
    }

    public GenericTableBuilder WithDateFormat(DateFormat format)
    {
        _dateFormat = format;
        return this;
    }

    public GenericTableBuilder WithNumberFormat(string columnName, NumberFormat format)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty", nameof(columnName));

        _numberFormats ??= new();
        _numberFormats[columnName] = format;
        return this;
    }

    public GenericTableBuilder WithConditionalStyle(string columnName, Func<CellValue, bool> condition, Action<CellStyleBuilder> style)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty", nameof(columnName));
        ArgumentNullException.ThrowIfNull(condition);
        ArgumentNullException.ThrowIfNull(style);

        _conditionalStyles ??= new();
        _conditionalStyles[columnName] = (condition, style);
        return this;
    }

    public WorkSheet Build()
    {
        var sheet = new WorkSheet(_sheetName);
        
        if (_table.ColumnCount == 0 || _table.RowCount == 0)
            return sheet;

        var headers = _table.Headers.ToList();
        headers = FilterAndOrderColumns(headers);
        
        if (headers.Count == 0)
            return sheet;

        AddHeaders(sheet, headers);
        AddDataRows(sheet, headers);

        if (!_autoFitColumns) return sheet;
        for (uint i = 0; i < headers.Count; i++)
            switch (_autoFitCalibration)
            {
                case null:
                    sheet.AutoFitColumn(i);
                    break;
                default:
                    sheet.AutoFitColumn(i, _autoFitCalibration.Value, _autoFitBaseLine);
                    break;
            }

        return sheet;
    }

    private List<string> FilterAndOrderColumns(List<string> headers)
    {
        var filtered = headers;

        if (_includeColumns != null)
            filtered = filtered.Where(h => _includeColumns.Contains(h)).ToList();

        if (_excludeColumns != null)
            filtered = filtered.Where(h => !_excludeColumns.Contains(h)).ToList();

        if (_columnOrder == null) 
            return filtered;
        
        var ordered = _columnOrder.Where(filtered.Contains).ToList();
        ordered.AddRange(filtered.Where(c => !_columnOrder.Contains(c)));

        return ordered;
    }

    private void AddHeaders(WorkSheet sheet, List<string> headers)
    {
        for (var i = 0; i < headers.Count; i++)
            if (_headerStyleAction != null)
                sheet.AddCell(new((uint)i, 0), new(headers[i]), builder => builder.WithStyle(_headerStyleAction));
            else
                sheet.AddCell(new((uint)i, 0), new(headers[i]), null);
    }

    private void AddDataRows(WorkSheet sheet, List<string> headers)
    {
        for (var rowIndex = 0; rowIndex < _table.RowCount; rowIndex++)
        for (var colIndex = 0; colIndex < headers.Count; colIndex++)
        {
            var headerName = headers[colIndex];
            var originalColIndex = _table.GetColumnIndex(headerName);
            var cellValue = _table.GetValue(originalColIndex, rowIndex);

            if (cellValue != null)
                AddCellValue(sheet, (uint)colIndex, (uint)(rowIndex + 1), cellValue, headerName);
        }
    }

    private void AddCellValue(WorkSheet sheet, uint col, uint row, CellValue cellValue, string? propertyName)
    {
        if (_trimWhitespace && cellValue.IsString())
            cellValue = new(cellValue.Value.AsT2.Trim());

        if (_columnParsers != null && propertyName != null && _columnParsers.TryGetValue(propertyName, out var parser))
            cellValue = parser(cellValue);

        sheet.AddCell(new(col, row), cellValue, builder =>
        {
            if (_preserveOriginalValue)
            {
                var originalValue = cellValue.Value.Match<string>(
                    d => d.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    l => l.ToString(),
                    s => s,
                    dt => dt.ToString("O"),
                    dto => dto.ToString("O"),
                    f => f.Value
                );
                builder.WithMetadata(metadata => metadata.WithOriginalValue(originalValue));
            }

            if (cellValue.IsDateTime() && _dateFormat != null)
                builder.WithStyle(style => style.WithFormatCode(GetDateFormatCode(_dateFormat.Value)));

            if (cellValue.IsDecimal() && _numberFormats != null && propertyName != null && _numberFormats.TryGetValue(propertyName, out var numberFormat))
                builder.WithStyle(style => style.WithFormatCode(GetNumberFormatCode(numberFormat)));

            if (_conditionalStyles == null || propertyName == null ||
                !_conditionalStyles.TryGetValue(propertyName, out var conditionalStyle)) return;
            if (conditionalStyle.condition(cellValue))
                builder.WithStyle(conditionalStyle.style);
        });
    }

    private static string GetDateFormatCode(DateFormat format) => format switch
    {
        DateFormat.IsoDateTime => "yyyy-mm-dd hh:mm:ss",
        DateFormat.IsoDate => "yyyy-mm-dd",
        DateFormat.DateTime => "dd/mm/yyyy hh:mm:ss",
        DateFormat.DateOnly => "dd/mm/yyyy",
        DateFormat.TimeOnly => "hh:mm:ss",
        _ => "yyyy-mm-dd hh:mm:ss"
    };

    private static string GetNumberFormatCode(NumberFormat format) => format switch
    {
        NumberFormat.Integer => "0",
        NumberFormat.Float2 => "0.00",
        NumberFormat.Float3 => "0.000",
        NumberFormat.Float4 => "0.0000",
        _ => "0.00"
    };
}
