using System.Text.Json;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public static class WorksheetBuilder
{
    public static GenericTableBuilder FromJson(string jsonContent)
    {
        var jsonDoc = JsonDocument.Parse(jsonContent);
        var jsonRoot = jsonDoc.RootElement;
        
        if (jsonRoot.ValueKind != JsonValueKind.Array && jsonRoot.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("JSON must be an array or object", nameof(jsonContent));
        
        var table = JsonToGenericTableConverter.Convert(jsonRoot, trimWhitespace: false);
        return GenericTableBuilder.FromGenericTable(table).WithPreserveOriginalValue(true).WithTrimWhitespace(true).WithSheetName("Sheet1");
    }

    public static GenericTableBuilder FromJsonFile(string filePath)
    {
        var jsonContent = File.ReadAllText(filePath);
        return FromJson(jsonContent);
    }

    public static GenericTableBuilder FromGenericTable(GenericTable table) => GenericTableBuilder.FromGenericTable(table);

    public static GenericTableBuilder FromCsv(string csvContent, bool hasHeader = true)
    {
        var table = CsvToGenericTableConverter.Convert(csvContent, hasHeader);
        return GenericTableBuilder.FromGenericTable(table);
    }

    public static GenericTableBuilder FromCsvFile(string filePath, bool hasHeader = true)
    {
        var csvContent = File.ReadAllText(filePath);
        return FromCsv(csvContent, hasHeader);
    }
}
