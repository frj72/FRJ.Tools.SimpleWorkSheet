using FRJ.Tools.SimpleWorkSheet.Components.Import;

namespace FRJ.Tools.SimpleWorksheetTests;

public class JsonWorkbookBuilderTests
{
    [Fact]
    public void FromJson_ValidJsonArray_ReturnsBuilder()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void FromJson_ValidObject_ReturnsBuilder()
    {
        const string json = """{"a": 1, "b": 2}""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void Build_CreatesWorkbookWithDefaultName()
    {
        const string json = """[{"name": "John"}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json).Build();
        
        Assert.NotNull(workbook);
        Assert.Equal("Workbook", workbook.Name);
    }

    [Fact]
    public void Build_CreatesWorkbookWithSingleSheet()
    {
        const string json = """[{"name": "John"}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json).Build();
        
        Assert.Single(workbook.Sheets);
        Assert.Equal("Data", workbook.Sheets.First().Name);
    }

    [Fact]
    public void WithWorkbookName_SetsCustomName()
    {
        const string json = """[{"test": "value"}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json)
            .WithWorkbookName("CustomWorkbook")
            .Build();
        
        Assert.Equal("CustomWorkbook", workbook.Name);
    }

    [Fact]
    public void WithWorkbookName_EmptyString_ThrowsException()
    {
        const string json = """[{"test": "value"}]""";
        var builder = JsonWorkbookBuilder.FromJson(json);
        
        Assert.Throws<ArgumentException>(() => builder.WithWorkbookName(""));
    }

    [Fact]
    public void WithDataSheetName_SetsCustomSheetName()
    {
        const string json = """[{"test": "value"}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json)
            .WithDataSheetName("CustomSheet")
            .Build();
        
        Assert.Equal("CustomSheet", workbook.Sheets.First().Name);
    }

    [Fact]
    public void WithDataSheetName_EmptyString_ThrowsException()
    {
        const string json = """[{"test": "value"}]""";
        var builder = JsonWorkbookBuilder.FromJson(json);
        
        Assert.Throws<ArgumentException>(() => builder.WithDataSheetName(""));
    }

    [Fact]
    public void GetColumnIndexByName_ExistingColumn_ReturnsIndex()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        var index = builder.GetColumnIndexByName("age");
        
        Assert.NotNull(index);
        Assert.True(index >= 0);
    }

    [Fact]
    public void GetColumnIndexByName_NonExistingColumn_ReturnsNull()
    {
        const string json = """[{"name": "John"}]""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        var index = builder.GetColumnIndexByName("nonexistent");
        
        Assert.Null(index);
    }

    [Fact]
    public void GetColumnRangeByName_ExistingColumn_ReturnsRange()
    {
        const string json = """[{"name": "John"}, {"name": "Jane"}]""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        var range = builder.GetColumnRangeByName("name");
        
        Assert.NotNull(range);
        Assert.Equal(1u, range.Value.From.Y);
        Assert.Equal(2u, range.Value.To.Y);
    }

    [Fact]
    public void GetColumnRangeByName_NonExistingColumn_ReturnsNull()
    {
        const string json = """[{"name": "John"}]""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        var range = builder.GetColumnRangeByName("nonexistent");
        
        Assert.Null(range);
    }

    [Fact]
    public void Build_ChainedMethods_WorksCorrectly()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json)
            .WithWorkbookName("People")
            .WithDataSheetName("PersonData")
            .WithTrimWhitespace(true)
            .AutoFitAllColumns()
            .Build();
        
        Assert.Equal("People", workbook.Name);
        Assert.Equal("PersonData", workbook.Sheets.First().Name);
        Assert.NotEmpty(workbook.Sheets.First().ExplicitColumnWidths);
    }

    [Fact]
    public void GetColumnIndexByName_NestedProperty_FindsColumn()
    {
        const string json = """[{"user": {"name": "John"}}]""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        var index = builder.GetColumnIndexByName("user.name");
        
        Assert.NotNull(index);
        Assert.Equal(0, index.Value);
    }

    [Fact]
    public void GetColumnRangeByName_NestedProperty_ReturnsCorrectRange()
    {
        const string json = """[{"user": {"age": 30}}, {"user": {"age": 25}}]""";
        
        var builder = JsonWorkbookBuilder.FromJson(json);
        var range = builder.GetColumnRangeByName("user.age");
        
        Assert.NotNull(range);
        Assert.Equal(0u, range.Value.From.X);
        Assert.Equal(0u, range.Value.To.X);
        Assert.Equal(1u, range.Value.From.Y);
        Assert.Equal(2u, range.Value.To.Y);
    }

    [Fact]
    public void WithPreserveOriginalValue_True_PreservesInMetadata()
    {
        const string json = """[{"price": 19.99}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json)
            .WithPreserveOriginalValue(true)
            .Build();
        
        var sheet = workbook.Sheets.First();
        var cell = sheet.Cells.Cells[new(0, 1)];
        Assert.NotNull(cell.Metadata);
        Assert.NotNull(cell.Metadata.OriginalValue);
        Assert.Equal("19.99", cell.Metadata.OriginalValue);
    }

    [Fact]
    public void WithPreserveOriginalValue_False_DoesNotPreserve()
    {
        const string json = """[{"price": 19.99}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json)
            .WithPreserveOriginalValue(false)
            .Build();
        
        var sheet = workbook.Sheets.First();
        var cell = sheet.Cells.Cells[new(0, 1)];
        Assert.Null(cell.Metadata);
    }

    [Fact]
    public void WithColumnParser_TransformsValues()
    {
        const string json = """[{"price": 100}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json)
            .WithColumnParser("price", value => new(value.Value.AsT0 * 2))
            .Build();
        
        var sheet = workbook.Sheets.First();
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.Equal(200m, value.Value.AsT0);
    }

    [Fact]
    public void WithColumnParser_MultipleColumns_AllTransformed()
    {
        const string json = """[{"price": 100, "tax": 10}]""";
        
        var workbook = JsonWorkbookBuilder.FromJson(json)
            .WithColumnParser("price", value => new(value.Value.AsT0 * 1.5m))
            .WithColumnParser("tax", value => new(value.Value.AsT0 * 2))
            .Build();
        
        var sheet = workbook.Sheets.First();
        Assert.Equal(150m, sheet.GetValue(0, 1)?.Value.AsT0);
        Assert.Equal(20m, sheet.GetValue(1, 1)?.Value.AsT0);
    }
}
