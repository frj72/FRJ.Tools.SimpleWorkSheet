using System.Globalization;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public static class CsvToGenericTableConverter
{
    public static GenericTable Convert(string csvContent, bool hasHeader = true)
    {
        var data = CsvParser.Parse(csvContent, hasHeader, out var headers);
        
        if (headers == null || headers.Count == 0)
            return GenericTable.Create();

        var table = new GenericTable([..headers]);
        
        foreach (var row in data)
        {
            var cellValues = new List<CellValue?>();
            
            for (var i = 0; i < headers.Count; i++)
                if (i < row.Count && !string.IsNullOrWhiteSpace(row[i]))
                {
                    var value = row[i];
                    cellValues.Add(ParseCellValue(value));
                }
                else
                    cellValues.Add(null);

            table.AddRow(cellValues);
        }

        return table;
    }

    public static CellValue ParseCellValue(string value)
    {
        if (DateTime.TryParse(value, out var dateValue))
            return new(dateValue);
        
        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var decimalValue))
            return new(decimalValue);
        
        if (long.TryParse(value, out var longValue))
            return new(longValue);
        
        return new(value);
    }
}
