using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class JsonArrayImportExample : IExample
{
    public string Name => "JSON Array Import";
    public string Description => "Imports a JSON array into a worksheet (one row per object)";

    public void Run()
    {
        var sheet = new WorkSheet("JsonImport");

       var jsonPath = Path.Combine("Resources", "Data", "Json", "users.json");
        var json = File.ReadAllText(jsonPath);

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Array) throw new ArgumentException("Expected JSON array", nameof(json));

        var headers = new List<string>();
        foreach (var p in root.EnumerateArray().SelectMany(el => el.EnumerateObject().Where(p => !headers.Contains(p.Name)))) headers.Add(p.Name);

        var baseOptions = ImportOptionsBuilder.Create()
            .WithSourceIdentifier("json")
            .WithPreserveOriginalValue(true)
            .WithTrimWhitespace(true)
            .Build();

        for (uint c = 0; c < headers.Count; c++)
        {
            var header = headers[(int)c];
            sheet.AddCell(c, 0, header, configure: cell => cell
                .WithColor("4472C4")
                .WithFont(f => f.Bold().WithColor("FFFFFF"))
                .FromImportedValue(header, baseOptions));
        }

        var nameToIndex = headers
            .Select((h, i) => new { h, i })
            .ToDictionary(x => x.h, x => x.i);

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

        ExampleRunner.SaveWorkSheet(sheet, "29_JsonArrayImport.xlsx");
    }
}
