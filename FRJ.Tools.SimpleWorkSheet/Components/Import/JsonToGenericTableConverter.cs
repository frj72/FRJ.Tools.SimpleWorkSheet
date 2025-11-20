using System.Globalization;
using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public static class JsonToGenericTableConverter
{
    public static GenericTable Convert(JsonElement jsonRoot, bool trimWhitespace = true)
    {
        return jsonRoot.ValueKind switch
        {
            JsonValueKind.Array => ConvertJsonArray(jsonRoot, trimWhitespace),
            JsonValueKind.Object => ConvertJsonObject(jsonRoot, trimWhitespace),
            _ => GenericTable.Create()
        };
    }

    public static GenericTable ConvertJsonArray(JsonElement jsonArray, bool trimWhitespace = true)
    {
        if (jsonArray.GetArrayLength() == 0)
            return GenericTable.Create();

        var propertyNames = DiscoverSchema(jsonArray);
        var table = new GenericTable([..propertyNames]);

        foreach (var item in jsonArray.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.Object)) table.AddRow(propertyNames.Select(propertyName => GetNestedProperty(item, propertyName)).Select(propertyValue => propertyValue.HasValue ? ConvertJsonValue(propertyValue.Value, trimWhitespace) : null).ToList());

        return table;
    }

    public static GenericTable ConvertJsonObject(JsonElement jsonObject, bool trimWhitespace = true)
    {
        var properties = jsonObject.EnumerateObject().Select(p => p.Name).ToList();
        var table = new GenericTable([..properties]);
        table.AddRow(properties.Select(propertyName => GetNestedProperty(jsonObject, propertyName)).Select(propertyValue => propertyValue.HasValue ? ConvertJsonValue(propertyValue.Value, trimWhitespace) : null).ToList());
        return table;
    }

    public static List<string> DiscoverSchema(JsonElement jsonArray)
    {
        var propertyNames = new HashSet<string>();

        foreach (var item in jsonArray.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.Object))
            FlattenProperties(item, "", propertyNames);

        return [..propertyNames];
    }

    public static void FlattenProperties(JsonElement element, string prefix, HashSet<string> propertyNames)
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

    public static JsonElement? GetNestedProperty(JsonElement element, string propertyPath)
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

    public static CellValue? ConvertJsonValue(JsonElement jsonValue, bool trimWhitespace = true)
    {
        switch (jsonValue.ValueKind)
        {
            case JsonValueKind.String:
                var stringValue = jsonValue.GetString();
                if (stringValue == null)
                    return null;
                
                if (trimWhitespace)
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
