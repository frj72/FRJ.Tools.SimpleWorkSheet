namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

internal static class CsvParser
{
    public static List<List<string>> Parse(string csvContent, bool hasHeader, out List<string>? headers)
    {
        var lines = csvContent.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        
        if (lines.Length == 0)
        {
            headers = null;
            return [];
        }

        var startIndex = 0;
        
        if (hasHeader)
        {
            headers = ParseLine(lines[0]);
            startIndex = 1;
        }
        else
        {
            var firstRow = ParseLine(lines[0]);
            headers = Enumerable.Range(1, firstRow.Count).Select(i => $"Column{i}").ToList();
        }

        var data = new List<List<string>>();
        
        for (var i = startIndex; i < lines.Length; i++)
            data.Add(ParseLine(lines[i]));

        return data;
    }

    private static List<string> ParseLine(string line)
    {
        var fields = new List<string>();
        var currentField = "";
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];

            switch (c)
            {
                case '"' when inQuotes && i + 1 < line.Length && line[i + 1] == '"':
                    currentField += '"';
                    i++;
                    break;
                case '"':
                    inQuotes = !inQuotes;
                    break;
                case ',' when !inQuotes:
                    fields.Add(currentField);
                    currentField = "";
                    break;
                default:
                    currentField += c;
                    break;
            }
        }

        fields.Add(currentField);
        return fields;
    }
}
