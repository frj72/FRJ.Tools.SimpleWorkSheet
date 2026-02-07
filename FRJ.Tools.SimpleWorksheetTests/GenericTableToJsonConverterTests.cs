using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class GenericTableToJsonConverterTests
{
    [Fact]
    public void ToJson_EmptyTable_ReturnsEmptyArray()
    {
        var table = GenericTable.Create();

        var json = GenericTableToJsonConverter.ToJson(table);

        Assert.NotNull(json);
        var jsonDoc = JsonDocument.Parse(json);
        Assert.Equal(JsonValueKind.Array, jsonDoc.RootElement.ValueKind);
        Assert.Equal(0, jsonDoc.RootElement.GetArrayLength());
    }

    [Fact]
    public void ToJson_SingleRowTable_ReturnsArrayWithOneObject()
    {
        var table = GenericTable.Create("Name", "Age");
        table.AddRow(new CellValue("Alice"), new CellValue(28));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        Assert.Equal(1, jsonDoc.RootElement.GetArrayLength());
        var firstRow = jsonDoc.RootElement[0];
        Assert.Equal("Alice", firstRow.GetProperty("Name").GetString());
        Assert.Equal(28, firstRow.GetProperty("Age").GetDecimal());
    }

    [Fact]
    public void ToJson_MultipleRows_ReturnsArrayWithMultipleObjects()
    {
        var table = GenericTable.Create("Name", "Age");
        table.AddRow(new CellValue("Alice"), new CellValue(28));
        table.AddRow(new CellValue("Bob"), new CellValue(35));
        table.AddRow(new CellValue("Charlie"), new CellValue(42));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        Assert.Equal(3, jsonDoc.RootElement.GetArrayLength());
        Assert.Equal("Bob", jsonDoc.RootElement[1].GetProperty("Name").GetString());
        Assert.Equal(35, jsonDoc.RootElement[1].GetProperty("Age").GetDecimal());
    }

    [Fact]
    public void ToJson_NullValues_WritesJsonNull()
    {
        var table = GenericTable.Create("Name", "Salary");
        table.AddRow(new CellValue("Alice"), null);
        table.AddRow(new CellValue("Bob"), new CellValue(50000m));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        Assert.Equal(JsonValueKind.Null, jsonDoc.RootElement[0].GetProperty("Salary").ValueKind);
        Assert.Equal(50000m, jsonDoc.RootElement[1].GetProperty("Salary").GetDecimal());
    }

    [Fact]
    public void ToJson_DecimalValues_WritesAsNumbers()
    {
        var table = GenericTable.Create("Price", "Tax");
        table.AddRow(new CellValue(99.99m), new CellValue(8.50m));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        Assert.Equal(JsonValueKind.Number, row.GetProperty("Price").ValueKind);
        Assert.Equal(99.99m, row.GetProperty("Price").GetDecimal());
        Assert.Equal(8.50m, row.GetProperty("Tax").GetDecimal());
    }

    [Fact]
    public void ToJson_LongValues_WritesAsNumbers()
    {
        var table = GenericTable.Create("ID", "Count");
        table.AddRow(new CellValue(1001L), new CellValue(5000L));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        Assert.Equal(JsonValueKind.Number, row.GetProperty("ID").ValueKind);
        Assert.Equal(1001, row.GetProperty("ID").GetInt64());
        Assert.Equal(5000, row.GetProperty("Count").GetInt64());
    }

    [Fact]
    public void ToJson_StringValues_WritesAsStrings()
    {
        var table = GenericTable.Create("FirstName", "LastName");
        table.AddRow(new CellValue("Alice"), new CellValue("Smith"));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        Assert.Equal(JsonValueKind.String, row.GetProperty("FirstName").ValueKind);
        Assert.Equal("Alice", row.GetProperty("FirstName").GetString());
        Assert.Equal("Smith", row.GetProperty("LastName").GetString());
    }

    [Fact]
    public void ToJson_DateTimeValues_WritesAsIso8601Strings()
    {
        var date = new DateTime(2025, 1, 15, 10, 30, 45, DateTimeKind.Utc);
        var table = GenericTable.Create("HireDate");
        table.AddRow(new CellValue(date));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        var dateString = row.GetProperty("HireDate").GetString();
        Assert.NotNull(dateString);
        Assert.Contains("2025-01-15", dateString);
        Assert.True(DateTime.TryParse(dateString, out var parsed));
        Assert.Equal(date.Year, parsed.Year);
        Assert.Equal(date.Month, parsed.Month);
        Assert.Equal(date.Day, parsed.Day);
    }

    [Fact]
    public void ToJson_DateTimeOffsetValues_WritesAsIso8601Strings()
    {
        var dateOffset = new DateTimeOffset(2025, 3, 20, 14, 15, 30, TimeSpan.FromHours(-5));
        var table = GenericTable.Create("Timestamp");
        table.AddRow(new CellValue(dateOffset));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        var dateString = row.GetProperty("Timestamp").GetString();
        Assert.NotNull(dateString);
        Assert.True(DateTimeOffset.TryParse(dateString, out var parsed));
        Assert.Equal(dateOffset.Year, parsed.Year);
        Assert.Equal(dateOffset.Offset, parsed.Offset);
    }

    [Fact]
    public void ToJson_FormulaValues_WritesFormulaTextAsString()
    {
        var table = GenericTable.Create("Formula");
        table.AddRow(new CellValue(new CellFormula("=SUM(A1:A10)")));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        Assert.Equal("=SUM(A1:A10)", row.GetProperty("Formula").GetString());
    }

    [Fact]
    public void ToJson_MixedDataTypes_WritesCorrectly()
    {
        var table = GenericTable.Create("Name", "ID", "Salary", "HireDate", "Active");
        table.AddRow(
            new CellValue("Alice"),
            new CellValue(1001L),
            new CellValue(75000.50m),
            new CellValue(new DateTime(2020, 1, 15)),
            new CellValue("TRUE")
        );

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        Assert.Equal("Alice", row.GetProperty("Name").GetString());
        Assert.Equal(1001, row.GetProperty("ID").GetInt64());
        Assert.Equal(75000.50m, row.GetProperty("Salary").GetDecimal());
        Assert.NotNull(row.GetProperty("HireDate").GetString());
        Assert.Equal("TRUE", row.GetProperty("Active").GetString());
    }

    [Fact]
    public void ToJson_PreservesColumnOrder()
    {
        var table = GenericTable.Create("Z", "A", "M");
        table.AddRow(new CellValue("z1"), new CellValue("a1"), new CellValue("m1"));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        var properties = row.EnumerateObject().Select(p => p.Name).ToList();
        Assert.Equal(["Z", "A", "M"], properties);
    }

    [Fact]
    public void ToJson_EmptyStrings_PreservesEmptyStrings()
    {
        var table = GenericTable.Create("Name", "Value");
        table.AddRow(new CellValue(""), new CellValue("test"));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        var row = jsonDoc.RootElement[0];
        Assert.Equal("", row.GetProperty("Name").GetString());
        Assert.NotEqual(JsonValueKind.Null, row.GetProperty("Name").ValueKind);
    }

    [Fact]
    public void ToJsonFile_CreatesValidFile()
    {
        var table = GenericTable.Create("Name", "Age");
        table.AddRow(new CellValue("Alice"), new CellValue(28));
        var tempFile = Path.GetTempFileName();

        try
        {
            GenericTableToJsonConverter.ToJsonFile(table, tempFile);

            Assert.True(File.Exists(tempFile));
            var json = File.ReadAllText(tempFile);
            var jsonDoc = JsonDocument.Parse(json);
            Assert.Equal(1, jsonDoc.RootElement.GetArrayLength());
            Assert.Equal("Alice", jsonDoc.RootElement[0].GetProperty("Name").GetString());
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Fact]
    public void ToJson_WithStream_WritesValidJson()
    {
        var table = GenericTable.Create("Name", "Age");
        table.AddRow(new CellValue("Bob"), new CellValue(35));

        using var stream = new MemoryStream();
        GenericTableToJsonConverter.ToJson(table, stream);
        stream.Position = 0;
        using var reader = new StreamReader(stream);
        var json = reader.ReadToEnd();

        var jsonDoc = JsonDocument.Parse(json);
        Assert.Equal(1, jsonDoc.RootElement.GetArrayLength());
        Assert.Equal("Bob", jsonDoc.RootElement[0].GetProperty("Name").GetString());
    }

    [Fact]
    public void ToJson_RoundTrip_PreservesData()
    {
        var originalTable = GenericTable.Create("Name", "Age", "Salary");
        originalTable.AddRow(new CellValue("Alice"), new CellValue(28), new CellValue(75000m));
        originalTable.AddRow(new CellValue("Bob"), new CellValue(35), null);
        originalTable.AddRow(new CellValue("Charlie"), new CellValue(42), new CellValue(85000m));

        var json = GenericTableToJsonConverter.ToJson(originalTable);
        var jsonDoc = JsonDocument.Parse(json);
        var importedTable = JsonToGenericTableConverter.Convert(jsonDoc.RootElement);

        Assert.Equal(originalTable.ColumnCount, importedTable.ColumnCount);
        Assert.Equal(originalTable.RowCount, importedTable.RowCount);
        Assert.Equal(originalTable.Headers, importedTable.Headers);

        for (var row = 0; row < originalTable.RowCount; row++)
            for (var col = 0; col < originalTable.ColumnCount; col++)
            {
                var originalValue = originalTable.GetValue(col, row);
                var importedValue = importedTable.GetValue(col, row);

                if (originalValue == null)
                    Assert.Null(importedValue);
                else if (originalValue.IsString())
                {
                    Assert.NotNull(importedValue);
                    Assert.True(importedValue.IsString());
                    Assert.Equal(originalValue.Value.AsT2, importedValue.Value.AsT2);
                }
                else if (originalValue.IsDecimal())
                {
                    Assert.NotNull(importedValue);
                    Assert.True(importedValue.IsDecimal());
                    Assert.Equal(originalValue.Value.AsT0, importedValue.Value.AsT0);
                }
            }
    }

    [Fact]
    public void ToJson_LargeTable_PerformsWell()
    {
        var table = GenericTable.Create("ID", "Name", "Value");
        for (var i = 0; i < 1000; i++)
            table.AddRow(new CellValue(i), new CellValue($"Name{i}"), new CellValue(i * 100m));

        var json = GenericTableToJsonConverter.ToJson(table);

        var jsonDoc = JsonDocument.Parse(json);
        Assert.Equal(1000, jsonDoc.RootElement.GetArrayLength());
        Assert.Equal("Name500", jsonDoc.RootElement[500].GetProperty("Name").GetString());
        Assert.Equal(50000m, jsonDoc.RootElement[500].GetProperty("Value").GetDecimal());
    }
}
