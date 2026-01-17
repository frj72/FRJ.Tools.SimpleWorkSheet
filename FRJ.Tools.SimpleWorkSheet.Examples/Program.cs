using System.Reflection;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples;

public static class Program
{
    public static void Main()
    {
        var examples = Assembly.GetExecutingAssembly()
            .GetTypes()
            .Where(t => typeof(IExample).IsAssignableFrom(t) && t is { IsInterface: false, IsAbstract: false })
            .Where(t => ExampleRegistry.Numbers.ContainsKey(t))
            .Select(t => Activator.CreateInstance(t, ExampleRegistry.Numbers[t]))
            .OfType<IExample>()
            .OrderBy(e => e.ExampleNumber)
            .ToList();

        Console.WriteLine("FRJ.Tools.SimpleWorkSheet - Examples");
        Console.WriteLine("====================================\n");

        RunAllExamples(examples);
    }

    private static void RunAllExamples(List<IExample> examples)
    {
        Console.WriteLine("Running all examples...\n");
        
        foreach (var example in examples) ExampleRunner.RunExample(example);

        Console.WriteLine($"\n✓ All {examples.Count} examples completed!");
        Console.WriteLine($"Output files are in: {Path.Combine(Directory.GetCurrentDirectory(), "Output")}");
    }
}
