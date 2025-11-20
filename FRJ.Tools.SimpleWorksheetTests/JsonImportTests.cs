using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class JsonImportTests
{
    [Fact]
    public void JsonArrayImport_BuildsHeadersAndRows_WithTyping_AndMetadata()
    {
        const string json = """
                            [
                              { "firstName": "Ben", "lastName": "Jensen", "age": 30, "salary": 1200.00 },
                              { "firstName": "Petra", "lastName": "Hansen", "age": 40, "salary": 3200.00 }
                            ]
                            """;

        var sheet = new WorkSheet("Test");
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var headers = new List<string>();
        foreach (var p in root.EnumerateArray().SelectMany(el => el.EnumerateObject().Where(p => !headers.Contains(p.Name)))) headers.Add(p.Name);

        var baseOptions = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("json")
            .WithPreserveOriginalValue(true)
            .Build();

        for (uint c = 0; c < headers.Count; c++)
        {
            var c1 = c;
            sheet.AddCell(c, 0, headers[(int)c], configure: cell => cell.FromImportedValue(headers[(int)c1], baseOptions));
        }

        var nameToIndex = headers.Select((h, i) => new { h, i }).ToDictionary(x => x.h, x => x.i);
        var optionsBuilder = ImportOptionsBuilder.FromOptions(baseOptions);
        if (nameToIndex.TryGetValue("age", out var ageIdx))
            optionsBuilder = optionsBuilder.WithColumnParser(ageIdx, s => new(int.Parse(s)));
        if (nameToIndex.TryGetValue("salary", out var salIdx))
            optionsBuilder = optionsBuilder.WithColumnParser(salIdx, s => new(decimal.Parse(s)));
        var options = optionsBuilder.Build();

        uint r = 1;
        foreach (var el in root.EnumerateArray())
        {
            for (uint c = 0; c < headers.Count; c++)
            {
                var key = headers[(int)c];
                var raw = el.TryGetProperty(key, out var prop) ? prop.ToString() : string.Empty;
                var processed = raw.ProcessRawValue(options);
                var value = options.ColumnParsers != null
                            && options.ColumnParsers.TryGetValue((int)c, out var parser)
                    ? parser(processed)
                    : new(processed);

                sheet.AddCell(c, r, value, configure: cell => cell.FromImportedValue(raw, options));
            }

            r++;
        }

        Assert.Equal("firstName", sheet.GetValue(0, 0)?.AsString());
        Assert.Equal("lastName", sheet.GetValue(1, 0)?.AsString());
        Assert.Equal("age", sheet.GetValue(2, 0)?.AsString());
        Assert.Equal("salary", sheet.GetValue(3, 0)?.AsString());

        Assert.Equal("Ben", sheet.GetValue(0, 1)?.AsString());
        Assert.Equal("Jensen", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal(30, sheet.GetValue(2, 1)?.AsInt());
        Assert.Equal(1200.00m, (sheet.GetValue(3, 1) ?? throw new InvalidOperationException()).AsDecimal());

        Assert.Equal("Petra", sheet.GetValue(0, 2)?.AsString());
        Assert.Equal("Hansen", sheet.GetValue(1, 2)?.AsString());
        Assert.Equal(40, sheet.GetValue(2, 2)?.AsInt());
        Assert.Equal(3200.00m, (sheet.GetValue(3, 2) ?? new CellValue(0)).AsDecimal());

        var firstDataCell = sheet.Cells.Cells[new(0, 1)];
        Assert.Equal("json", firstDataCell.Metadata?.Source);
        Assert.Equal("Ben", firstDataCell.Metadata?.OriginalValue);
    }

    [Fact]
    public void JsonArrayImport_WithMissingKeys_FillsEmptyCells()
    {
        const string json = """
                            [
                              { "firstName": "Ben", "lastName": "Jensen", "age": 30, "salary": 1200.00 },
                              { "firstName": "Petra", "lastName": "Hansen", "age": 40, "salary": 3200.00 },
                              { "firstName": "Peter", "lastName": "Jazz", "age": 43, "alive": true },
                              { "firstName": "Mary", "middleName": "Jean", "lastName": "Johnson", "age": 66, "alive": true, "salary": 1200 }
                            ]
                            """;

        var sheet = new WorkSheet("Test");
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var headers = new List<string>();
        foreach (var p in root.EnumerateArray().SelectMany(el => el.EnumerateObject().Where(p => !headers.Contains(p.Name)))) headers.Add(p.Name);

        var baseOptions = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("json")
            .WithPreserveOriginalValue(true)
            .Build();

        for (uint c = 0; c < headers.Count; c++)
        {
            var c1 = c;
            sheet.AddCell(c, 0, headers[(int)c], configure: cell => cell.FromImportedValue(headers[(int)c1], baseOptions));
        }

        var nameToIndex = headers.Select((h, i) => new { h, i }).ToDictionary(x => x.h, x => x.i);
        var optionsBuilder = ImportOptionsBuilder.FromOptions(baseOptions);
        if (nameToIndex.TryGetValue("age", out var ageIdx))
            optionsBuilder = optionsBuilder.WithColumnParser(ageIdx, s => new(int.Parse(s)));
        if (nameToIndex.TryGetValue("salary", out var salIdx))
            optionsBuilder = optionsBuilder.WithColumnParser(salIdx, s => string.IsNullOrEmpty(s) ? new("") : new CellValue(decimal.Parse(s)));
        if (nameToIndex.TryGetValue("alive", out var aliveIdx))
            optionsBuilder = optionsBuilder.WithColumnParser(aliveIdx, s => string.IsNullOrEmpty(s) ? new("") : new CellValue(bool.Parse(s).ToString()));
        var options = optionsBuilder.Build();

        uint r = 1;
        foreach (var el in root.EnumerateArray())
        {
            for (uint c = 0; c < headers.Count; c++)
            {
                var key = headers[(int)c];
                var raw = el.TryGetProperty(key, out var prop) ? prop.ToString() : string.Empty;
                var processed = raw.ProcessRawValue(options);
                var value = options.ColumnParsers != null
                            && options.ColumnParsers.TryGetValue((int)c, out var parser)
                    ? parser(processed)
                    : new(processed);

                sheet.AddCell(c, r, value, configure: cell => cell.FromImportedValue(raw, options));
            }

            r++;
        }

        Assert.Equal(6, headers.Count);
        Assert.Equal("firstName", headers[0]);
        Assert.Equal("lastName", headers[1]);
        Assert.Equal("age", headers[2]);
        Assert.Equal("salary", headers[3]);
        Assert.Equal("alive", headers[4]);
        Assert.Equal("middleName", headers[5]);

        Assert.Equal("Ben", sheet.GetValue(0, 1)?.AsString());
        Assert.Equal("Jensen", sheet.GetValue(1, 1)?.AsString());
        Assert.Equal(30, sheet.GetValue(2, 1)?.AsInt());
        Assert.Equal(1200.00m, (sheet.GetValue(3, 1) ?? throw new InvalidOperationException()).AsDecimal());
        Assert.Equal("", sheet.GetValue(4, 1)?.AsString());
        Assert.Equal("", sheet.GetValue(5, 1)?.AsString());

        Assert.Equal("Petra", sheet.GetValue(0, 2)?.AsString());
        Assert.Equal("Hansen", sheet.GetValue(1, 2)?.AsString());
        Assert.Equal(40, sheet.GetValue(2, 2)?.AsInt());
        Assert.Equal(3200.00m, (sheet.GetValue(3, 2) ?? throw new InvalidOperationException()).AsDecimal());
        Assert.Equal("", sheet.GetValue(4, 2)?.AsString());
        Assert.Equal("", sheet.GetValue(5, 2)?.AsString());

        Assert.Equal("Peter", sheet.GetValue(0, 3)?.AsString());
        Assert.Equal("Jazz", sheet.GetValue(1, 3)?.AsString());
        Assert.Equal(43, sheet.GetValue(2, 3)?.AsInt());
        Assert.Equal("", sheet.GetValue(3, 3)?.AsString());
        Assert.Equal("True", sheet.GetValue(4, 3)?.AsString());
        Assert.Equal("", sheet.GetValue(5, 3)?.AsString());

        Assert.Equal("Mary", sheet.GetValue(0, 4)?.AsString());
        Assert.Equal("Johnson", sheet.GetValue(1, 4)?.AsString());
        Assert.Equal(66, sheet.GetValue(2, 4)?.AsInt());
        Assert.Equal(1200m, (sheet.GetValue(3, 4) ?? throw new InvalidOperationException()).AsDecimal());
        Assert.Equal("True", sheet.GetValue(4, 4)?.AsString());
        Assert.Equal("Jean", sheet.GetValue(5, 4)?.AsString());

        var firstDataCell = sheet.Cells.Cells[new(0, 1)];
        Assert.Equal("json", firstDataCell.Metadata?.Source);
        Assert.Equal("Ben", firstDataCell.Metadata?.OriginalValue);
    }
}
