using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class WorkbookBuilder
{
    private string _workbookName = "Workbook";
    private readonly WorksheetBuilder _worksheetBuilder;
    private readonly List<Func<WorkSheet, WorkSheet>> _chartBuilders = [];
    private bool _freezeHeaderRow;

    internal string DataSheetName { get; private set; } = "Data";

    private WorkbookBuilder(JsonElement jsonRoot)
    {
        if (jsonRoot.ValueKind != JsonValueKind.Array && jsonRoot.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("JSON must be an array or object", nameof(jsonRoot));
        
        _worksheetBuilder = WorksheetBuilder.FromJson(jsonRoot.GetRawText()).WithSheetName(DataSheetName);
    }

    public static WorkbookBuilder FromJson(string jsonContent)
    {
        var jsonDoc = JsonDocument.Parse(jsonContent);
        return new(jsonDoc.RootElement);
    }

    public static WorkbookBuilder FromJsonFile(string filePath)
    {
        var jsonContent = File.ReadAllText(filePath);
        return FromJson(jsonContent);
    }

    public WorkbookBuilder WithWorkbookName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Workbook name cannot be empty", nameof(name));
        
        _workbookName = name;
        return this;
    }

    public WorkbookBuilder WithDataSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Data sheet name cannot be empty", nameof(name));
        
        DataSheetName = name;
        _worksheetBuilder.WithSheetName(name);
        return this;
    }

    public WorkbookBuilder WithPreserveOriginalValue(bool preserve)
    {
        _worksheetBuilder.WithPreserveOriginalValue(preserve);
        return this;
    }

    public WorkbookBuilder WithTrimWhitespace(bool trim)
    {
        _worksheetBuilder.WithTrimWhitespace(trim);
        return this;
    }

    public WorkbookBuilder WithHeaderStyle(Action<CellStyleBuilder> styleAction)
    {
        _worksheetBuilder.WithHeaderStyle(styleAction);
        return this;
    }

    public WorkbookBuilder WithColumnParser(string columnName, Func<CellValue, CellValue> parser)
    {
        _worksheetBuilder.WithColumnParser(columnName, parser);
        return this;
    }

    public WorkbookBuilder AutoFitAllColumns()
    {
        _worksheetBuilder.AutoFitAllColumns();
        return this;
    }

    public WorkbookBuilder WithFreezeHeaderRow()
    {
        _freezeHeaderRow = true;
        return this;
    }

    public WorkbookBuilder WithChart(Action<ChartBuilder> chartConfig)
    {
        ArgumentNullException.ThrowIfNull(chartConfig);

        var builder = new ChartBuilder(this);
        chartConfig(builder);
        _chartBuilders.Add(builder.Build);
        
        return this;
    }

    public WorkBook Build()
    {
        var dataSheet = _worksheetBuilder.Build();
        
        if (_freezeHeaderRow)
            dataSheet.FreezePanes(1, 0);
        
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

