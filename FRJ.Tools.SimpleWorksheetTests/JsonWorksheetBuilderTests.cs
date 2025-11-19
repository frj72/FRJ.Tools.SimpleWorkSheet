using FRJ.Tools.SimpleWorkSheet.Components.Import;

namespace FRJ.Tools.SimpleWorksheetTests;

public class JsonWorksheetBuilderTests
{
    [Fact]
    public void FromJson_ValidJsonArray_ReturnsBuilder()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var builder = JsonWorksheetBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void FromJson_ValidObject_ReturnsBuilder()
    {
        const string json = """{"valid": "this is a flat object"}""";
        
        var builder = JsonWorksheetBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void FromJson_EmptyArray_ReturnsBuilder()
    {
        const string json = "[]";
        
        var builder = JsonWorksheetBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void Build_EmptyArray_ReturnsEmptyWorksheet()
    {
        const string json = "[]";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.NotNull(sheet);
        Assert.Equal("Sheet1", sheet.Name);
        Assert.Empty(sheet.Cells.Cells);
    }

    [Fact]
    public void Build_SimpleObject_CreatesHeadersAndData()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.NotNull(sheet);
        Assert.Equal(4, sheet.Cells.Cells.Count);
        
        var headerName = sheet.GetValue(0, 0);
        var headerAge = sheet.GetValue(1, 0);
        Assert.NotNull(headerName);
        Assert.NotNull(headerAge);
        Assert.Equal("name", headerName.Value.AsT2);
        Assert.Equal("age", headerAge.Value.AsT2);
        
        var dataName = sheet.GetValue(0, 1);
        var dataAge = sheet.GetValue(1, 1);
        Assert.NotNull(dataName);
        Assert.NotNull(dataAge);
        Assert.Equal("John", dataName.Value.AsT2);
        Assert.Equal(30m, dataAge.Value.AsT0);
    }

    [Fact]
    public void Build_MultipleObjects_CreatesMultipleRows()
    {
        const string json = """[{"name": "John"}, {"name": "Jane"}, {"name": "Bob"}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.Equal(4, sheet.Cells.Cells.Count);
        Assert.Equal("name", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("John", sheet.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal("Jane", sheet.GetValue(0, 2)?.Value.AsT2);
        Assert.Equal("Bob", sheet.GetValue(0, 3)?.Value.AsT2);
    }

    [Fact]
    public void Build_DetectsStringType()
    {
        const string json = """[{"text": "hello"}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.True(value.IsString());
        Assert.Equal("hello", value.Value.AsT2);
    }

    [Fact]
    public void Build_DetectsNumberType()
    {
        const string json = """[{"count": 42}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.True(value.IsDecimal());
        Assert.Equal(42m, value.Value.AsT0);
    }

    [Fact]
    public void Build_DetectsDecimalNumber()
    {
        const string json = """[{"price": 19.99}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.True(value.IsDecimal());
        Assert.Equal(19.99m, value.Value.AsT0);
    }

    [Fact]
    public void Build_DetectsDateString()
    {
        const string json = """[{"date": "2025-01-15"}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.True(value.IsDateTime());
        Assert.Equal(new(2025, 1, 15), value.Value.AsT3);
    }

    [Fact]
    public void Build_DetectsBooleanTrue()
    {
        const string json = """[{"active": true}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.True(value.IsString());
        Assert.Equal("TRUE", value.Value.AsT2);
    }

    [Fact]
    public void Build_DetectsBooleanFalse()
    {
        const string json = """[{"active": false}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.True(value.IsString());
        Assert.Equal("FALSE", value.Value.AsT2);
    }

    [Fact]
    public void Build_SkipsNullValues()
    {
        const string json = """[{"name": "John", "age": null}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.Equal(3, sheet.Cells.Cells.Count);
        Assert.NotNull(sheet.GetValue(0, 0));
        Assert.NotNull(sheet.GetValue(1, 0));
        Assert.NotNull(sheet.GetValue(0, 1));
        Assert.Null(sheet.GetValue(1, 1));
    }

    [Fact]
    public void WithSheetName_SetsCustomName()
    {
        const string json = """[{"test": "value"}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithSheetName("CustomSheet")
            .Build();
        
        Assert.Equal("CustomSheet", sheet.Name);
    }

    [Fact]
    public void WithSheetName_EmptyString_ThrowsException()
    {
        const string json = """[{"test": "value"}]""";
        var builder = JsonWorksheetBuilder.FromJson(json);
        
        Assert.Throws<ArgumentException>(() => builder.WithSheetName(""));
    }

    [Fact]
    public void WithSheetName_Whitespace_ThrowsException()
    {
        const string json = """[{"test": "value"}]""";
        var builder = JsonWorksheetBuilder.FromJson(json);
        
        Assert.Throws<ArgumentException>(() => builder.WithSheetName("   "));
    }

    [Fact]
    public void WithPreserveOriginalValue_True_StoresOriginalInMetadata()
    {
        const string json = """[{"price": 19.99}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithPreserveOriginalValue(true)
            .Build();
        
        var cell = sheet.Cells.Cells[new(0, 1)];
        Assert.NotNull(cell.Metadata);
        Assert.NotNull(cell.Metadata.OriginalValue);
        Assert.Equal("19.99", cell.Metadata.OriginalValue);
    }

    [Fact]
    public void WithPreserveOriginalValue_False_DoesNotStoreMetadata()
    {
        const string json = """[{"price": 19.99}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithPreserveOriginalValue(false)
            .Build();
        
        var cell = sheet.Cells.Cells[new(0, 1)];
        Assert.Null(cell.Metadata);
    }

    [Fact]
    public void WithTrimWhitespace_True_TrimsStrings()
    {
        const string json = """[{"name": "  John  "}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithTrimWhitespace(true)
            .Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.Equal("John", value.Value.AsT2);
    }

    [Fact]
    public void WithTrimWhitespace_False_PreservesWhitespace()
    {
        const string json = """[{"name": "  John  "}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithTrimWhitespace(false)
            .Build();
        
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.Equal("  John  ", value.Value.AsT2);
    }

    [Fact]
    public void Build_DiscoversSchemaDynamically()
    {
        const string json = """[{"a": 1, "b": 2}, {"a": 3, "c": 4}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        var headers = new List<string?>
        {
            sheet.GetValue(0, 0)?.Value.AsT2,
            sheet.GetValue(1, 0)?.Value.AsT2,
            sheet.GetValue(2, 0)?.Value.AsT2
        };
        
        Assert.Contains("a", headers);
        Assert.Contains("b", headers);
        Assert.Contains("c", headers);
    }

    [Fact]
    public void Build_HandlesMissingProperties()
    {
        const string json = """[{"name": "John", "age": 30}, {"name": "Jane"}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.NotNull(sheet.GetValue(0, 1));
        Assert.NotNull(sheet.GetValue(1, 1));
        Assert.NotNull(sheet.GetValue(0, 2));
        Assert.Null(sheet.GetValue(1, 2));
    }

    [Fact]
    public void Build_MultipleProperties_CreatesCorrectStructure()
    {
        const string json = """[{"id": 1, "name": "John", "age": 30, "active": true}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.Equal(8, sheet.Cells.Cells.Count);
        Assert.Equal(1m, sheet.GetValue(0, 1)?.Value.AsT0);
        Assert.Equal("John", sheet.GetValue(1, 1)?.Value.AsT2);
        Assert.Equal(30m, sheet.GetValue(2, 1)?.Value.AsT0);
        Assert.Equal("TRUE", sheet.GetValue(3, 1)?.Value.AsT2);
    }

    [Fact]
    public void Build_ChainedMethods_WorksCorrectly()
    {
        const string json = """[{"name": "  Test  ", "price": 99.99}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithSheetName("Products")
            .WithTrimWhitespace(true)
            .WithPreserveOriginalValue(false)
            .Build();
        
        Assert.Equal("Products", sheet.Name);
        Assert.Equal("Test", sheet.GetValue(0, 1)?.Value.AsT2);
        
        var cell = sheet.Cells.Cells[new(1, 1)];
        Assert.Null(cell.Metadata);
    }

    [Fact]
    public void FromJson_FlatObject_ReturnsBuilder()
    {
        const string json = """{"price_1": 0.12, "price_2": 0.25, "price_3": 0.1}""";
        
        var builder = JsonWorksheetBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void Build_FlatObject_CreatesHeadersAndSingleRow()
    {
        const string json = """{"name": "John", "age": 30, "city": "NYC"}""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.NotNull(sheet);
        Assert.Equal(6, sheet.Cells.Cells.Count);
        
        Assert.Equal("name", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("age", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("city", sheet.GetValue(2, 0)?.Value.AsT2);
        
        Assert.Equal("John", sheet.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal(30m, sheet.GetValue(1, 1)?.Value.AsT0);
        Assert.Equal("NYC", sheet.GetValue(2, 1)?.Value.AsT2);
    }

    [Fact]
    public void Build_FlatObjectWithNumbers_ConvertsCorrectly()
    {
        const string json = """{"price_1": 0.12, "price_2": 0.25, "price_3": 0.1}""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.Equal(6, sheet.Cells.Cells.Count);
        Assert.Equal("price_1", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal(0.12m, sheet.GetValue(0, 1)?.Value.AsT0);
        Assert.Equal(0.25m, sheet.GetValue(1, 1)?.Value.AsT0);
        Assert.Equal(0.1m, sheet.GetValue(2, 1)?.Value.AsT0);
    }

    [Fact]
    public void Build_FlatObject_WithPreserveOriginalValue()
    {
        const string json = """{"value": 42.5}""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithPreserveOriginalValue(true)
            .Build();
        
        var cell = sheet.Cells.Cells[new(0, 1)];
        Assert.NotNull(cell.Metadata);
        Assert.NotNull(cell.Metadata.OriginalValue);
        Assert.Equal("42.5", cell.Metadata.OriginalValue);
    }

    [Fact]
    public void Build_EmptyObject_ReturnsEmptyWorksheet()
    {
        const string json = "{}";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.NotNull(sheet);
        Assert.Empty(sheet.Cells.Cells);
    }

    [Fact]
    public void FromJson_InvalidJsonType_ThrowsException()
    {
        const string json = "123";
        
        Assert.Throws<ArgumentException>(() => JsonWorksheetBuilder.FromJson(json));
    }

    [Fact]
    public void WithHeaderStyle_AppliesStylesToHeaders()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithHeaderStyle(style => style.WithFillColor("FF0000").WithFont(f => f.Bold()))
            .Build();
        
        var headerCell = sheet.Cells.Cells[new(0, 0)];
        Assert.Equal("FF0000", headerCell.Style.FillColor);
        Assert.True(headerCell.Style.Font?.Bold);
    }

    [Fact]
    public void WithColumnParser_TransformsColumnValues()
    {
        const string json = """[{"price": 100}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithColumnParser("price", value => new(value.Value.AsT0 * 1.1m))
            .Build();
        
        var dataCell = sheet.GetValue(0, 1);
        Assert.NotNull(dataCell);
        Assert.Equal(110m, dataCell.Value.AsT0);
    }

    [Fact]
    public void WithColumnParser_EmptyColumnName_ThrowsException()
    {
        const string json = """[{"test": "value"}]""";
        var builder = JsonWorksheetBuilder.FromJson(json);
        
        Assert.Throws<ArgumentException>(() => builder.WithColumnParser("", v => v));
    }

    [Fact]
    public void WithColumnParser_MultipleColumns()
    {
        const string json = """[{"price": 100, "quantity": 5}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithColumnParser("price", value => new(value.Value.AsT0 * 2))
            .WithColumnParser("quantity", value => new(value.Value.AsT0 * 3))
            .Build();
        
        Assert.Equal(200m, sheet.GetValue(0, 1)?.Value.AsT0);
        Assert.Equal(15m, sheet.GetValue(1, 1)?.Value.AsT0);
    }

    [Fact]
    public void AutoFitAllColumns_SetsColumnWidths()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .AutoFitAllColumns()
            .Build();
        
        Assert.NotEmpty(sheet.ExplicitColumnWidths);
        Assert.True(sheet.ExplicitColumnWidths.Count >= 2);
    }

    [Fact]
    public void AutoFitAllColumns_WorksWithFlatObject()
    {
        const string json = """{"a": 1, "b": 2, "c": 3}""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .AutoFitAllColumns()
            .Build();
        
        Assert.NotEmpty(sheet.ExplicitColumnWidths);
        Assert.True(sheet.ExplicitColumnWidths.Count >= 3);
    }

    [Fact]
    public void Build_CombinedFeatures_WorksTogether()
    {
        const string json = """[{"price": 100, "name": "Item"}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json)
            .WithSheetName("Products")
            .WithHeaderStyle(style => style.WithFillColor("4472C4"))
            .WithColumnParser("price", value => new(value.Value.AsT0 * 1.2m))
            .AutoFitAllColumns()
            .Build();
        
        Assert.Equal("Products", sheet.Name);
        Assert.Equal("4472C4", sheet.Cells.Cells[new(0, 0)].Style.FillColor);
        Assert.Equal(120m, sheet.GetValue(0, 1)?.Value.AsT0);
        Assert.NotEmpty(sheet.ExplicitColumnWidths);
    }

    [Fact]
    public void Build_SparseData_HandlesCorrectly()
    {
        const string json = """[{"name": "John", "age": 30}, {"name": "Jane"}, {"age": 25, "city": "NYC"}]""";
        
        var sheet = JsonWorksheetBuilder.FromJson(json).Build();
        
        Assert.Equal("John", sheet.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal(30m, sheet.GetValue(1, 1)?.Value.AsT0);
        Assert.Equal("Jane", sheet.GetValue(0, 2)?.Value.AsT2);
        Assert.Null(sheet.GetValue(1, 2));
        Assert.Null(sheet.GetValue(0, 3));
        Assert.Equal(25m, sheet.GetValue(1, 3)?.Value.AsT0);
        Assert.Equal("NYC", sheet.GetValue(2, 3)?.Value.AsT2);
    }
}
