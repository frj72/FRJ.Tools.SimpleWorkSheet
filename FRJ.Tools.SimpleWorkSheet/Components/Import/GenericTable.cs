using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public class GenericTable
{
    private readonly List<string> _headers;
    private readonly List<List<CellValue?>> _rows;

    public GenericTable()
    {
        _headers = [];
        _rows = [];
    }

    public GenericTable(List<string> headers)
    {
        ArgumentNullException.ThrowIfNull(headers);
        _headers = headers;
        _rows = [];
    }

    public IReadOnlyList<string> Headers => _headers.AsReadOnly();
    public IReadOnlyList<IReadOnlyList<CellValue?>> Rows => _rows.Select(r => r.AsReadOnly()).ToList().AsReadOnly();
    public int ColumnCount => _headers.Count;
    public int RowCount => _rows.Count;

    public void AddHeader(string header)
    {
        if (string.IsNullOrWhiteSpace(header))
            throw new ArgumentException("Header cannot be empty", nameof(header));
        
        _headers.Add(header);
    }

    public void AddHeaders(params string[] headers)
    {
        ArgumentNullException.ThrowIfNull(headers);
        
        foreach (var header in headers)
            AddHeader(header);
    }

    public void AddRow(params CellValue?[] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        
        if (values.Length != _headers.Count && _headers.Count > 0)
            throw new ArgumentException($"Row must have {_headers.Count} values to match headers", nameof(values));
        
        _rows.Add([..values]);
    }

    public void AddRow(List<CellValue?> values)
    {
        ArgumentNullException.ThrowIfNull(values);
        
        if (values.Count != _headers.Count && _headers.Count > 0)
            throw new ArgumentException($"Row must have {_headers.Count} values to match headers", nameof(values));
        
        _rows.Add([..values]);
    }

    public CellValue? GetValue(int columnIndex, int rowIndex)
    {
        if (columnIndex < 0 || columnIndex >= _headers.Count)
            throw new ArgumentOutOfRangeException(nameof(columnIndex));
        
        if (rowIndex < 0 || rowIndex >= _rows.Count)
            throw new ArgumentOutOfRangeException(nameof(rowIndex));
        
        return _rows[rowIndex][columnIndex];
    }

    public CellValue? GetValue(string columnName, int rowIndex)
    {
        var columnIndex = _headers.IndexOf(columnName);
        return columnIndex == -1 ? throw new ArgumentException($"Column '{columnName}' not found", nameof(columnName)) : GetValue(columnIndex, rowIndex);
    }

    public string GetHeader(int columnIndex)
    {
        if (columnIndex < 0 || columnIndex >= _headers.Count)
            throw new ArgumentOutOfRangeException(nameof(columnIndex));
        
        return _headers[columnIndex];
    }

    public int GetColumnIndex(string columnName)
    {
        var index = _headers.IndexOf(columnName);
        return index == -1 ? throw new ArgumentException($"Column '{columnName}' not found", nameof(columnName)) : index;
    }

    public bool HasColumn(string columnName) => _headers.Contains(columnName);

    public static GenericTable Create() => new();

    public static GenericTable Create(params string[] headers) => new([..headers]);
}
