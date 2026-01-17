namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

public interface IExample
{
    string Name { get; }
    string Description { get; }
    int ExampleNumber { get; }
    void Run();
}
