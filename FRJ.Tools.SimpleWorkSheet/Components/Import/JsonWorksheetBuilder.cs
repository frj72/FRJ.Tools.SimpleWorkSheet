using System.Globalization;
using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class JsonWorksheetBuilder
{
    private readonly JsonElement _jsonRoot;
    private string _sheetName = "Sheet1";
    private bool _preserveOriginalValue = true;
    private bool _trimWhitespace = true;
    private Action<CellStyleBuilder>? _headerStyleAction;
    private Dictionary<string, Func<CellValue, CellValue>>? _columnParsers;
    private bool _autoFitColumns;

    private JsonWorksheetBuilder(JsonElement jsonRoot)
    {
        if (jsonRoot.ValueKind != JsonValueKind.Array && jsonRoot.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("JSON must be an array or object", nameof(jsonRoot));
        
        _jsonRoot = jsonRoot;
    }

    public static JsonWorksheetBuilder FromJson(string jsonContent)
    {
        var jsonDoc = JsonDocument.Parse(jsonContent);
        return new(jsonDoc.RootElement);
    }

    public static JsonWorksheetBuilder FromJsonFile(string filePath)
    {
        var jsonContent = File.ReadAllText(filePath);
        return FromJson(jsonContent);
    }

    public JsonWorksheetBuilder WithSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet name cannot be empty", nameof(name));
        
        _sheetName = name;
        return this;
    }

    public JsonWorksheetBuilder WithPreserveOriginalValue(bool preserve)
    {
        _preserveOriginalValue = preserve;
        return this;
    }

    public JsonWorksheetBuilder WithTrimWhitespace(bool trim)
    {
        _trimWhitespace = trim;
        return this;
    }

    public JsonWorksheetBuilder WithHeaderStyle(Action<CellStyleBuilder> styleAction)
    {
        _headerStyleAction = styleAction ?? throw new ArgumentNullException(nameof(styleAction));
        return this;
    }

    public JsonWorksheetBuilder WithColumnParser(string columnName, Func<CellValue, CellValue> parser)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty", nameof(columnName));

        _columnParsers ??= new();
        _columnParsers[columnName] = parser ?? throw new ArgumentNullException(nameof(parser));
        return this;
    }

    public JsonWorksheetBuilder AutoFitAllColumns()
    {
        _autoFitColumns = true;
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

        if (_preserveOriginalValue)
        {
            var originalValue = propertyValue.GetRawText();
            sheet.AddCell(new(col, row), cellValue, builder =>
            {
                builder.WithMetadata(metadata => metadata.WithOriginalValue(originalValue));
            });
        }
        else
            sheet.AddCell(new(col, row), cellValue);
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
}
