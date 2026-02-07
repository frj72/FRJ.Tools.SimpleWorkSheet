using System.Reflection;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ExampleMappingTests
{
    [Fact]
    public void ExampleFilesCsv_AllDiscoveredExamplesArePresent()
    {
        var csvPath = Path.Combine(GetExamplesProjectPath(), "ExampleFiles.csv");
        Assert.True(File.Exists(csvPath), $"ExampleFiles.csv not found at {csvPath}");

        var csvEntries = File.ReadAllLines(csvPath)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line => line.Split(','))
            .Where(parts => parts.Length == 3)
            .Select(parts => parts[2].Trim())
            .ToHashSet();

        var examplesAssembly = Assembly.Load("FRJ.Tools.SimpleWorkSheet.Examples");
        var discoveredExamples = examplesAssembly.GetTypes()
            .Where(t => t.GetInterfaces().Any(i => i.Name == "IExample"))
            .Where(t => t is { IsInterface: false, IsAbstract: false })
            .Select(t => t.Name)
            .ToList();

        var missingFromCsv = discoveredExamples.Where(name => !csvEntries.Contains(name)).ToList();

        Assert.Empty(missingFromCsv);
    }

    [Fact]
    public void ExampleFilesCsv_AllEntriesHaveValidFormat()
    {
        var csvPath = Path.Combine(GetExamplesProjectPath(), "ExampleFiles.csv");
        var lines = File.ReadAllLines(csvPath)
            .Where(l => !string.IsNullOrWhiteSpace(l));

        foreach (var line in lines)
        {
            var parts = line.Split(',');
            Assert.Equal(3, parts.Length);

            var parsed = int.TryParse(parts[0], out var number);
            Assert.True(parsed, $"Invalid number in line: {line}");
            Assert.True(number is >= 1 and <= 116, $"Number out of range in line: {line}");

            Assert.True(parts[1].EndsWith(".xlsx"), $"Invalid filename in line: {line}");
            Assert.True(parts[1].StartsWith($"{number:000}_"), $"Filename doesn't match number in line: {line}");

            Assert.True(parts[2].EndsWith("Example"), $"Invalid class name in line: {line}");
        }
    }

    [Fact]
    public void ExampleFilesCsv_IsSortedByNumber()
    {
        var csvPath = Path.Combine(GetExamplesProjectPath(), "ExampleFiles.csv");
        var numbers = File.ReadAllLines(csvPath)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line => int.Parse(line.Split(',')[0]))
            .ToList();

        var sortedNumbers = numbers.OrderBy(n => n).ToList();
        Assert.Equal(sortedNumbers, numbers);
    }

    [Fact]
    public void ExampleFilesCsv_HasNoDuplicateNumbers()
    {
        var csvPath = Path.Combine(GetExamplesProjectPath(), "ExampleFiles.csv");
        var numbers = File.ReadAllLines(csvPath)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line => int.Parse(line.Split(',')[0]))
            .ToList();

        var duplicates = numbers.GroupBy(n => n).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        Assert.Empty(duplicates);
    }

    [Fact]
    public void ExampleFilesCsv_HasNoDuplicateClassNames()
    {
        var csvPath = Path.Combine(GetExamplesProjectPath(), "ExampleFiles.csv");
        var classNames = File.ReadAllLines(csvPath)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line => line.Split(',')[2].Trim())
            .ToList();

        var duplicates = classNames.GroupBy(n => n).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        Assert.Empty(duplicates);
    }

    private static string GetExamplesProjectPath()
    {
        var currentDir = Directory.GetCurrentDirectory();
        var projectRoot = currentDir;

        while (projectRoot != null && !Directory.Exists(Path.Combine(projectRoot, "FRJ.Tools.SimpleWorkSheet.Examples")))
            projectRoot = Directory.GetParent(projectRoot)?.FullName;

        return projectRoot == null 
            ? throw new DirectoryNotFoundException("Could not find FRJ.Tools.SimpleWorkSheet.Examples project") 
            : Path.Combine(projectRoot, "FRJ.Tools.SimpleWorkSheet.Examples");
    }
}
