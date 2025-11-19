using System.Globalization;
using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class WorksheetBuilder
{
    private readonly JsonElement _jsonRoot;
    private string _sheetName = "Sheet1";
    private bool _preserveOriginalValue = true;
    private bool _trimWhitespace = true;
    private Action<CellStyleBuilder>? _headerStyleAction;
    private Dictionary<string, Func<CellValue, CellValue>>? _columnParsers;
    private bool _autoFitColumns;
    private List<string>? _columnOrder;
    private HashSet<string>? _excludeColumns;
    private HashSet<string>? _includeColumns;
    private DateFormat? _dateFormat;
    private Dictionary<string, NumberFormat>? _numberFormats;
    private Dictionary<string, (Func<CellValue, bool> condition, Action<CellStyleBuilder> style)>? _conditionalStyles;

    private WorksheetBuilder(JsonElement jsonRoot)
    {
        if (jsonRoot.ValueKind != JsonValueKind.Array && jsonRoot.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("JSON must be an array or object", nameof(jsonRoot));
        
        _jsonRoot = jsonRoot;
    }

    public static WorksheetBuilder FromJson(string jsonContent)
    {
        var jsonDoc = JsonDocument.Parse(jsonContent);
        return new(jsonDoc.RootElement);
    }

    public static WorksheetBuilder FromJsonFile(string filePath)
    {
        var jsonContent = File.ReadAllText(filePath);
        return FromJson(jsonContent);
    }

    public static GenericTableBuilder FromGenericTable(GenericTable table) => GenericTableBuilder.FromGenericTable(table);

    public static WorksheetBuilder FromCsv(string csvContent, bool hasHeader = true)
    {
        var jsonContent = ConvertCsvToJson(csvContent, hasHeader);
        return FromJson(jsonContent);
    }

    public static WorksheetBuilder FromCsvFile(string filePath, bool hasHeader = true)
    {
        var csvContent = File.ReadAllText(filePath);
        return FromCsv(csvContent, hasHeader);
    }

    internal static string ConvertCsvToJson(string csvContent, bool hasHeader)
    {
        var data = CsvParser.Parse(csvContent, hasHeader, out var headers);
        
        if (headers == null || headers.Count == 0)
            return "[]";

        var jsonArray = new List<string>();
        
        foreach (var row in data)
        {
            var properties = new List<string>();
            for (var i = 0; i < Math.Min(headers.Count, row.Count); i++)
                properties.Add($"\"{headers[i]}\":\"{EscapeJsonString(row[i])}\"");
            
            jsonArray.Add("{" + string.Join(",", properties) + "}");
        }

        return "[" + string.Join(",", jsonArray) + "]";
    }

    private static string EscapeJsonString(string value) =>
        value.Replace("\\", @"\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");

    public WorksheetBuilder WithSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet name cannot be empty", nameof(name));
        
        _sheetName = name;
        return this;
    }

    public WorksheetBuilder WithPreserveOriginalValue(bool preserve)
    {
        _preserveOriginalValue = preserve;
        return this;
    }

    public WorksheetBuilder WithTrimWhitespace(bool trim)
    {
        _trimWhitespace = trim;
        return this;
    }

    public WorksheetBuilder WithHeaderStyle(Action<CellStyleBuilder> styleAction)
    {
        _headerStyleAction = styleAction ?? throw new ArgumentNullException(nameof(styleAction));
        return this;
    }

    public WorksheetBuilder WithColumnParser(string columnName, Func<CellValue, CellValue> parser)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty", nameof(columnName));

        _columnParsers ??= new();
        _columnParsers[columnName] = parser ?? throw new ArgumentNullException(nameof(parser));
        return this;
    }

    public WorksheetBuilder AutoFitAllColumns()
    {
        _autoFitColumns = true;
        return this;
    }

    public WorksheetBuilder WithColumnOrder(params string[] columnNames)
    {
        if (columnNames == null || columnNames.Length == 0)
            throw new ArgumentException("Column order must contain at least one column", nameof(columnNames));

        _columnOrder = [..columnNames];
        return this;
    }

    public WorksheetBuilder WithExcludeColumns(params string[] columnNames)
    {
        if (columnNames == null || columnNames.Length == 0)
            throw new ArgumentException("Exclude columns must contain at least one column", nameof(columnNames));

        _excludeColumns = [..columnNames];
        return this;
    }

    public WorksheetBuilder WithIncludeColumns(params string[] columnNames)
    {
        if (columnNames == null || columnNames.Length == 0)
            throw new ArgumentException("Include columns must contain at least one column", nameof(columnNames));

        _includeColumns = [..columnNames];
        return this;
    }

    public WorksheetBuilder WithDateFormat(DateFormat format)
    {
        _dateFormat = format;
        return this;
    }

    public WorksheetBuilder WithNumberFormat(string columnName, NumberFormat format)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty", nameof(columnName));

        _numberFormats ??= new();
        _numberFormats[columnName] = format;
        return this;
    }

    public WorksheetBuilder WithConditionalStyle(string columnName, Func<CellValue, bool> condition, Action<CellStyleBuilder> style)
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
        
        return _jsonRoot.ValueKind == JsonValueKind.Array ? BuildFromArray(sheet) : BuildFromObject(sheet);
    }

    private WorkSheet BuildFromArray(WorkSheet sheet)
    {
        if (_jsonRoot.GetArrayLength() == 0)
            return sheet;

        var properties = DiscoverSchemaFromArray(_jsonRoot);
        properties = FilterAndOrderColumns(properties);
        
        if (properties.Count == 0)
            return sheet;

        AddHeaders(sheet, properties);
        AddDataRowsFromArray(sheet, properties);

        if (!_autoFitColumns) return sheet;
        for (uint i = 0; i < properties.Count; i++)
            sheet.AutoFitColumn(i);

        return sheet;
    }

    private WorkSheet BuildFromObject(WorkSheet sheet)
    {
        var properties = _jsonRoot.EnumerateObject().Select(p => p.Name).ToList();
        properties = FilterAndOrderColumns(properties);
        
        if (properties.Count == 0)
            return sheet;

        AddHeaders(sheet, properties);
        AddDataRowFromObject(sheet, properties);

        if (!_autoFitColumns) return sheet;
        for (uint i = 0; i < properties.Count; i++)
            sheet.AutoFitColumn(i);

        return sheet;
    }

    private static List<string> DiscoverSchemaFromArray(JsonElement jsonArray)
    {
        var propertyNames = new HashSet<string>();

        foreach (var item in jsonArray.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.Object))
            FlattenProperties(item, "", propertyNames);

        return [..propertyNames];
    }

    private static void FlattenProperties(JsonElement element, string prefix, HashSet<string> propertyNames)
    {
        foreach (var property in element.EnumerateObject())
        {
            var propertyName = string.IsNullOrEmpty(prefix) ? property.Name : $"{prefix}.{property.Name}";

            if (property.Value.ValueKind == JsonValueKind.Object)
                FlattenProperties(property.Value, propertyName, propertyNames);
            else if (property.Value.ValueKind != JsonValueKind.Array)
                propertyNames.Add(propertyName);
        }
    }

    private List<string> FilterAndOrderColumns(List<string> properties)
    {
        var filtered = properties;

        if (_includeColumns != null)
            filtered = filtered.Where(p => _includeColumns.Contains(p)).ToList();

        if (_excludeColumns != null)
            filtered = filtered.Where(p => !_excludeColumns.Contains(p)).ToList();

        if (_columnOrder == null) 
            return filtered;
        
        var ordered = _columnOrder.Where(filtered.Contains).ToList();
        ordered.AddRange(filtered.Where(c => !_columnOrder.Contains(c)));

        return ordered;
    }

    private void AddHeaders(WorkSheet sheet, List<string> properties)
    {
        for (var i = 0; i < properties.Count; i++)
            if (_headerStyleAction != null)
                sheet.AddCell(new((uint)i, 0), new(properties[i]), builder => builder.WithStyle(_headerStyleAction));
            else
                sheet.AddCell(new((uint)i, 0), new CellValue(properties[i]));
    }

    private void AddDataRowsFromArray(WorkSheet sheet, List<string> properties)
    {
        uint rowIndex = 1;

        foreach (var item in _jsonRoot.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.Object))
        {
            for (var colIndex = 0; colIndex < properties.Count; colIndex++)
            {
                var propertyName = properties[colIndex];
                var propertyValue = GetNestedProperty(item, propertyName);

                if (propertyValue.HasValue)
                    AddCellValue(sheet, (uint)colIndex, rowIndex, propertyValue.Value, propertyName);
            }

            rowIndex++;
        }
    }

    private static JsonElement? GetNestedProperty(JsonElement element, string propertyPath)
    {
        var parts = propertyPath.Split('.');
        var current = element;

        foreach (var part in parts)
        {
            if (!current.TryGetProperty(part, out var next))
                return null;
            current = next;
        }

        return current;
    }

    private void AddDataRowFromObject(WorkSheet sheet, List<string> properties)
    {
        const uint rowIndex = 1;

        for (var colIndex = 0; colIndex < properties.Count; colIndex++)
        {
            var propertyName = properties[colIndex];
            var propertyValue = GetNestedProperty(_jsonRoot, propertyName);

            if (propertyValue.HasValue)
                AddCellValue(sheet, (uint)colIndex, rowIndex, propertyValue.Value, propertyName);
        }
    }

    private void AddCellValue(WorkSheet sheet, uint col, uint row, JsonElement propertyValue, string? propertyName = null)
    {
        var cellValue = ConvertJsonValueToCellValue(propertyValue);
        
        if (cellValue == null)
            return;

        if (_columnParsers != null && propertyName != null && _columnParsers.TryGetValue(propertyName, out var parser))
            cellValue = parser(cellValue);

        sheet.AddCell(new(col, row), cellValue, builder =>
        {
            if (_preserveOriginalValue)
                builder.WithMetadata(metadata => metadata.WithOriginalValue(propertyValue.GetRawText()));

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

    private CellValue? ConvertJsonValueToCellValue(JsonElement jsonValue)
    {
        switch (jsonValue.ValueKind)
        {
            case JsonValueKind.String:
                var stringValue = jsonValue.GetString();
                if (stringValue == null)
                    return null;
                
                if (_trimWhitespace)
                    stringValue = stringValue.Trim();

                if (DateTime.TryParse(stringValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateValue))
                    return new(dateValue);

                return new(stringValue);

            case JsonValueKind.Number:
                return new(jsonValue.GetDouble());

            case JsonValueKind.True:
                return new("TRUE");

            case JsonValueKind.False:
                return new("FALSE");

            case JsonValueKind.Null:

            case JsonValueKind.Undefined:
            case JsonValueKind.Object:
            case JsonValueKind.Array:
            default:
                return null;
        }
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
