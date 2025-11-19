using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class JsonWorkbookBuilder
{
    private string _workbookName = "Workbook";
    private readonly JsonWorksheetBuilder _worksheetBuilder;
    private readonly List<Func<WorkSheet, WorkSheet>> _chartBuilders = [];

    internal string DataSheetName { get; private set; } = "Data";

    private JsonWorkbookBuilder(JsonElement jsonRoot)
    {
        if (jsonRoot.ValueKind != JsonValueKind.Array && jsonRoot.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("JSON must be an array or object", nameof(jsonRoot));
        
        _worksheetBuilder = JsonWorksheetBuilder.FromJson(jsonRoot.GetRawText()).WithSheetName(DataSheetName);
    }

    public static JsonWorkbookBuilder FromJson(string jsonContent)
    {
        var jsonDoc = JsonDocument.Parse(jsonContent);
        return new(jsonDoc.RootElement);
    }

    public static JsonWorkbookBuilder FromJsonFile(string filePath)
    {
        var jsonContent = File.ReadAllText(filePath);
        return FromJson(jsonContent);
    }

    public JsonWorkbookBuilder WithWorkbookName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Workbook name cannot be empty", nameof(name));
        
        _workbookName = name;
        return this;
    }

    public JsonWorkbookBuilder WithDataSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Data sheet name cannot be empty", nameof(name));
        
        DataSheetName = name;
        _worksheetBuilder.WithSheetName(name);
        return this;
    }

    public JsonWorkbookBuilder WithPreserveOriginalValue(bool preserve)
    {
        _worksheetBuilder.WithPreserveOriginalValue(preserve);
        return this;
    }

    public JsonWorkbookBuilder WithTrimWhitespace(bool trim)
    {
        _worksheetBuilder.WithTrimWhitespace(trim);
        return this;
    }

    public JsonWorkbookBuilder WithHeaderStyle(Action<CellStyleBuilder> styleAction)
    {
        _worksheetBuilder.WithHeaderStyle(styleAction);
        return this;
    }

    public JsonWorkbookBuilder WithColumnParser(string columnName, Func<CellValue, CellValue> parser)
    {
        _worksheetBuilder.WithColumnParser(columnName, parser);
        return this;
    }

    public JsonWorkbookBuilder AutoFitAllColumns()
    {
        _worksheetBuilder.AutoFitAllColumns();
        return this;
    }

    public JsonWorkbookBuilder WithChart(Action<JsonChartBuilder> chartConfig)
    {
        ArgumentNullException.ThrowIfNull(chartConfig);

        var builder = new JsonChartBuilder(this);
        chartConfig(builder);
        _chartBuilders.Add(builder.Build);
        
        return this;
    }

    public WorkBook Build()
    {
        var dataSheet = _worksheetBuilder.Build();
        var sheets = new List<WorkSheet> { dataSheet };
        sheets.AddRange(_chartBuilders.Select(buildChartAction => buildChartAction(dataSheet)));

        return new(_workbookName, [..sheets]);
    }

    public int? GetColumnIndexByName(string columnName)
    {
        var dataSheet = _worksheetBuilder.Build();
        
        for (uint col = 0; col < 1000; col++)
        {
            var header = dataSheet.GetValue(col, 0);
            if (header?.Value.AsT2 == columnName)
                return (int)col;
        }

        return null;
    }

    public CellRange? GetColumnRangeByName(string columnName)
    {
        var colIndex = GetColumnIndexByName(columnName);
        if (colIndex == null)
            return null;

        var dataSheet = _worksheetBuilder.Build();
        var rowCount = CountDataRows(dataSheet);

        if (rowCount == 0)
            return null;

        return CellRange.FromBounds((uint)colIndex.Value, 1, (uint)colIndex.Value, rowCount);
    }

    private static uint CountDataRows(WorkSheet sheet)
    {
        uint maxRow = 0;
        
        foreach (var position in sheet.Cells.Cells.Keys.Where(position => position.Y > maxRow)) 
            maxRow = position.Y;

        return maxRow;
    }
}

