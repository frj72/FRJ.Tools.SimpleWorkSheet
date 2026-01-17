using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class WorkbookBuilder
{
    private string _workbookName = "Workbook";
    private readonly GenericTableBuilder _genericTableBuilder;
    private readonly List<Func<WorkSheet, WorkSheet>> _chartBuilders = [];
    private bool _freezeHeaderRow;

    internal string DataSheetName { get; private set; } = "Data";

    private WorkbookBuilder(GenericTable table)
    {
        ArgumentNullException.ThrowIfNull(table);
        _genericTableBuilder = GenericTableBuilder.FromGenericTable(table).WithSheetName(DataSheetName);
    }

    public static WorkbookBuilder FromJson(string jsonContent)
    {
        var jsonDoc = JsonDocument.Parse(jsonContent);
        var table = JsonToGenericTableConverter.Convert(jsonDoc.RootElement, trimWhitespace: true);
        return new(table);
    }

    public static WorkbookBuilder FromJsonFile(string filePath)
    {
        var jsonContent = File.ReadAllText(filePath);
        return FromJson(jsonContent);
    }

    public static WorkbookBuilder FromCsv(string csvContent, bool hasHeader = true)
    {
        var table = CsvToGenericTableConverter.Convert(csvContent, hasHeader);
        return FromGenericTable(table);
    }

    public static WorkbookBuilder FromCsvFile(string filePath, bool hasHeader = true)
    {
        var csvContent = File.ReadAllText(filePath);
        return FromCsv(csvContent, hasHeader);
    }

    public static WorkbookBuilder FromClass<T>(T source) where T:class
    {
        var jsonContent = JsonSerializer.Serialize(source);
        return FromJson(jsonContent);
    }

    public static WorkbookBuilder FromGenericTable(GenericTable table) => new(table);

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
        _genericTableBuilder.WithSheetName(name);
        return this;
    }

    public WorkbookBuilder WithPreserveOriginalValue(bool preserve)
    {
        _genericTableBuilder.WithPreserveOriginalValue(preserve);
        return this;
    }

    public WorkbookBuilder WithTrimWhitespace(bool trim)
    {
        _genericTableBuilder.WithTrimWhitespace(trim);
        return this;
    }

    public WorkbookBuilder WithHeaderStyle(Action<CellStyleBuilder> styleAction)
    {
        _genericTableBuilder.WithHeaderStyle(styleAction);
        return this;
    }

    public WorkbookBuilder WithColumnParser(string columnName, Func<CellValue, CellValue> parser)
    {
        _genericTableBuilder.WithColumnParser(columnName, parser);
        return this;
    }

    public WorkbookBuilder AutoFitAllColumns()
    {
        _genericTableBuilder.AutoFitAllColumns();
        return this;
    }

    public WorkbookBuilder AutoFitAllColumns(double calibration)
    {
        _genericTableBuilder.AutoFitAllColumns(calibration);
        return this;
    }

    public WorkbookBuilder WithColumnOrder(params string[] columnNames)
    {
        _genericTableBuilder.WithColumnOrder(columnNames);
        return this;
    }

    public WorkbookBuilder WithExcludeColumns(params string[] columnNames)
    {
        _genericTableBuilder.WithExcludeColumns(columnNames);
        return this;
    }

    public WorkbookBuilder WithIncludeColumns(params string[] columnNames)
    {
        _genericTableBuilder.WithIncludeColumns(columnNames);
        return this;
    }

    public WorkbookBuilder WithDateFormat(DateFormat format)
    {
        _genericTableBuilder.WithDateFormat(format);
        return this;
    }

    public WorkbookBuilder WithNumberFormat(string columnName, NumberFormat format)
    {
        _genericTableBuilder.WithNumberFormat(columnName, format);
        return this;
    }

    public WorkbookBuilder WithConditionalStyle(string columnName, Func<CellValue, bool> condition, Action<CellStyleBuilder> style)
    {
        _genericTableBuilder.WithConditionalStyle(columnName, condition, style);
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
        var dataSheet = _genericTableBuilder.Build();
        
        if (_freezeHeaderRow)
            dataSheet.FreezePanes(1, 0);
        
        var sheets = new List<WorkSheet> { dataSheet };
        sheets.AddRange(_chartBuilders.Select(buildChartAction => buildChartAction(dataSheet)));

        return new(_workbookName, [..sheets]);
    }

    public int? GetColumnIndexByName(string columnName)
    {
        var dataSheet = _genericTableBuilder.Build();
        
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

        var dataSheet = _genericTableBuilder.Build();
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

