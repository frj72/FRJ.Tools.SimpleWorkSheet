using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Import;

namespace FRJ.Tools.SimpleWorksheetTests;

public class JsonToGenericTableConverterTests
{
    [Fact]
    public void Convert_EmptyArray_ReturnsEmptyTable()
    {
        const string json = "[]";
        var jsonDoc = JsonDocument.Parse(json);
        
        var table = JsonToGenericTableConverter.Convert(jsonDoc.RootElement);
        
        Assert.NotNull(table);
        Assert.Equal(0, table.ColumnCount);
        Assert.Equal(0, table.RowCount);
    }

    [Fact]
    public void Convert_SimpleArray_CreatesTableWithHeaders()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var table = JsonToGenericTableConverter.Convert(jsonDoc.RootElement);
        
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(1, table.RowCount);
        Assert.Contains("name", table.Headers);
        Assert.Contains("age", table.Headers);
        Assert.Equal("John", table.GetValue("name", 0)?.Value.AsT2);
        Assert.Equal(30m, table.GetValue("age", 0)?.Value.AsT0);
    }

    [Fact]
    public void Convert_SimpleObject_CreatesSingleRowTable()
    {
        const string json = """{"name": "Alice", "score": 95.5}""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var table = JsonToGenericTableConverter.Convert(jsonDoc.RootElement);
        
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(1, table.RowCount);
        Assert.Equal("Alice", table.GetValue("name", 0)?.Value.AsT2);
        Assert.Equal(95.5m, table.GetValue("score", 0)?.Value.AsT0);
    }

    [Fact]
    public void ConvertJsonArray_MultipleRows_CreatesAllRows()
    {
        const string json = """[{"id": 1}, {"id": 2}, {"id": 3}]""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var table = JsonToGenericTableConverter.ConvertJsonArray(jsonDoc.RootElement);
        
        Assert.Equal(1, table.ColumnCount);
        Assert.Equal(3, table.RowCount);
        Assert.Equal(1m, table.GetValue(0, 0)?.Value.AsT0);
        Assert.Equal(2m, table.GetValue(0, 1)?.Value.AsT0);
        Assert.Equal(3m, table.GetValue(0, 2)?.Value.AsT0);
    }

    [Fact]
    public void ConvertJsonValue_String_ReturnsString()
    {
        const string json = "\"hello\"";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.ConvertJsonValue(jsonDoc.RootElement);
        
        Assert.NotNull(value);
        Assert.Equal("hello", value.Value.AsT2);
    }

    [Fact]
    public void ConvertJsonValue_Number_ReturnsDecimal()
    {
        const string json = "42.5";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.ConvertJsonValue(jsonDoc.RootElement);
        
        Assert.NotNull(value);
        Assert.Equal(42.5m, value.Value.AsT0);
    }

    [Fact]
    public void ConvertJsonValue_True_ReturnsString()
    {
        const string json = "true";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.ConvertJsonValue(jsonDoc.RootElement);
        
        Assert.NotNull(value);
        Assert.Equal("TRUE", value.Value.AsT2);
    }

    [Fact]
    public void ConvertJsonValue_False_ReturnsString()
    {
        const string json = "false";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.ConvertJsonValue(jsonDoc.RootElement);
        
        Assert.NotNull(value);
        Assert.Equal("FALSE", value.Value.AsT2);
    }

    [Fact]
    public void ConvertJsonValue_Null_ReturnsNull()
    {
        const string json = "null";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.ConvertJsonValue(jsonDoc.RootElement);
        
        Assert.Null(value);
    }

    [Fact]
    public void ConvertJsonValue_WithTrimWhitespace_TrimsString()
    {
        const string json = "\"  spaced  \"";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.ConvertJsonValue(jsonDoc.RootElement, trimWhitespace: true);
        
        Assert.NotNull(value);
        Assert.Equal("spaced", value.Value.AsT2);
    }

    [Fact]
    public void ConvertJsonValue_WithoutTrimWhitespace_PreservesSpaces()
    {
        const string json = "\"  spaced  \"";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.ConvertJsonValue(jsonDoc.RootElement, trimWhitespace: false);
        
        Assert.NotNull(value);
        Assert.Equal("  spaced  ", value.Value.AsT2);
    }

    [Fact]
    public void DiscoverSchema_SimpleArray_ReturnsAllProperties()
    {
        const string json = """[{"a": 1, "b": 2}, {"b": 3, "c": 4}]""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var schema = JsonToGenericTableConverter.DiscoverSchema(jsonDoc.RootElement);
        
        Assert.Equal(3, schema.Count);
        Assert.Contains("a", schema);
        Assert.Contains("b", schema);
        Assert.Contains("c", schema);
    }

    [Fact]
    public void DiscoverSchema_NestedObject_FlattensWithDotNotation()
    {
        const string json = """[{"user": {"name": "John", "age": 30}}]""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var schema = JsonToGenericTableConverter.DiscoverSchema(jsonDoc.RootElement);
        
        Assert.Equal(2, schema.Count);
        Assert.Contains("user.name", schema);
        Assert.Contains("user.age", schema);
    }

    [Fact]
    public void GetNestedProperty_SimplePath_ReturnsValue()
    {
        const string json = """{"name": "Alice"}""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.GetNestedProperty(jsonDoc.RootElement, "name");
        
        Assert.True(value.HasValue);
        Assert.Equal("Alice", value.Value.GetString());
    }

    [Fact]
    public void GetNestedProperty_NestedPath_ReturnsValue()
    {
        const string json = """{"user": {"profile": {"name": "Bob"}}}""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.GetNestedProperty(jsonDoc.RootElement, "user.profile.name");
        
        Assert.True(value.HasValue);
        Assert.Equal("Bob", value.Value.GetString());
    }

    [Fact]
    public void GetNestedProperty_NonExistentPath_ReturnsNull()
    {
        const string json = """{"name": "Test"}""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var value = JsonToGenericTableConverter.GetNestedProperty(jsonDoc.RootElement, "nonexistent");
        
        Assert.False(value.HasValue);
    }

    [Fact]
    public void Convert_NestedObjects_FlattensCorrectly()
    {
        const string json = """[{"person": {"name": "John", "age": 30}, "score": 85}]""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var table = JsonToGenericTableConverter.Convert(jsonDoc.RootElement);
        
        Assert.Equal(3, table.ColumnCount);
        Assert.Contains("person.name", table.Headers);
        Assert.Contains("person.age", table.Headers);
        Assert.Contains("score", table.Headers);
    }

    [Fact]
    public void Convert_MixedTypes_ConvertsCorrectly()
    {
        const string json = """[{"str": "text", "num": 42, "bool": true, "date": "2024-01-15"}]""";
        var jsonDoc = JsonDocument.Parse(json);
        
        var table = JsonToGenericTableConverter.Convert(jsonDoc.RootElement);
        
        Assert.Equal("text", table.GetValue("str", 0)?.Value.AsT2);
        Assert.Equal(42m, table.GetValue("num", 0)?.Value.AsT0);
        Assert.Equal("TRUE", table.GetValue("bool", 0)?.Value.AsT2);
    }

    [Fact]
    public void FlattenProperties_ThreeLevelsDeep_FlattensAll()
    {
        const string json = """{"a": {"b": {"c": 1}}}""";
        var jsonDoc = JsonDocument.Parse(json);
        var propertyNames = new HashSet<string>();
        
        JsonToGenericTableConverter.FlattenProperties(jsonDoc.RootElement, "", propertyNames);
        
        Assert.Single(propertyNames);
        Assert.Contains("a.b.c", propertyNames);
    }

    [Fact]
    public void ConvertJsonObject_EmptyObject_CreatesTableWithNoColumns()
    {
        const string json = "{}";
        var jsonDoc = JsonDocument.Parse(json);
        
        var table = JsonToGenericTableConverter.ConvertJsonObject(jsonDoc.RootElement);
        
        Assert.Equal(0, table.ColumnCount);
    }
}
