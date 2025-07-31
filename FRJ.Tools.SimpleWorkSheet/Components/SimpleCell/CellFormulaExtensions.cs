namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class CellFormulaExtensions
{
    public static CellFormula ToFormula(this string str) => new(str);
    public static string Evaluate(this CellFormula cellFormula, Func<string, string> f) => f(cellFormula.Value);
}
