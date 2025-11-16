namespace FRJ.Tools.SimpleWorkSheet.Showcase;

public interface IShowcase
{
    string Name { get; }
    string Description { get; }
    string Category { get; }
    void Run();
}
