namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class FormulaBuilder
{
    public static CellFormula VLookup(string lookupValue, string tableRange, int columnIndex, bool exactMatch = false)
    {
        var matchType = exactMatch ? "FALSE" : "TRUE";
        return new($"=VLOOKUP({lookupValue},{tableRange},{columnIndex},{matchType})");
    }

    public static CellFormula XLookup(string lookupValue, string lookupArray, string returnArray, string? ifNotFound)
    {
        var formula = ifNotFound != null
            ? $"=XLOOKUP({lookupValue},{lookupArray},{returnArray},{ifNotFound})"
            : $"=XLOOKUP({lookupValue},{lookupArray},{returnArray})";
        return new(formula);
    }

    public static CellFormula Index(string array, int rowNum, int? columnNum)
    {
        var formula = columnNum.HasValue
            ? $"=INDEX({array},{rowNum},{columnNum.Value})"
            : $"=INDEX({array},{rowNum})";
        return new(formula);
    }

    public static CellFormula Match(string lookupValue, string lookupArray, int matchType = 0) 
        => new($"=MATCH({lookupValue},{lookupArray},{matchType})");

    public static CellFormula IndexMatch(string returnArray, string lookupValue, string lookupArray, int matchType = 0) 
        => new($"=INDEX({returnArray},MATCH({lookupValue},{lookupArray},{matchType}))");

    public static CellFormula Concatenate(params string[] values) 
        => new($"=CONCATENATE({string.Join(",", values)})");

    public static CellFormula Concat(params string[] values) 
        => new($"=CONCAT({string.Join(",", values)})");

    public static CellFormula Left(string text, int numChars) 
        => new($"=LEFT({text},{numChars})");

    public static CellFormula Right(string text, int numChars) 
        => new($"=RIGHT({text},{numChars})");

    public static CellFormula Mid(string text, int startNum, int numChars) 
        => new($"=MID({text},{startNum},{numChars})");

    public static CellFormula Trim(string text) 
        => new($"=TRIM({text})");

    public static CellFormula Upper(string text) 
        => new($"=UPPER({text})");

    public static CellFormula Lower(string text) 
        => new($"=LOWER({text})");

    public static CellFormula Proper(string text) 
        => new($"=PROPER({text})");

    public static CellFormula Len(string text) 
        => new($"=LEN({text})");

    public static CellFormula If(string condition, string valueIfTrue, string valueIfFalse) 
        => new($"=IF({condition},{valueIfTrue},{valueIfFalse})");

    public static CellFormula And(params string[] conditions) 
        => new($"=AND({string.Join(",", conditions)})");

    public static CellFormula Or(params string[] conditions) 
        => new($"=OR({string.Join(",", conditions)})");

    public static CellFormula Not(string condition) 
        => new($"=NOT({condition})");

    public static CellFormula Xor(params string[] conditions) 
        => new($"=XOR({string.Join(",", conditions)})");

    public static CellFormula Today() 
        => new("=TODAY()");

    public static CellFormula Now() 
        => new("=NOW()");

    public static CellFormula Date(int year, int month, int day) 
        => new($"=DATE({year},{month},{day})");

    public static CellFormula Year(string date) 
        => new($"=YEAR({date})");

    public static CellFormula Month(string date) => new($"=MONTH({date})");

    public static CellFormula Day(string date) => new($"=DAY({date})");

    public static CellFormula DateDif(string startDate, string endDate, string unit) 
        => new($"=DATEDIF({startDate},{endDate},\"{unit}\")");

    public static CellFormula EoMonth(string startDate, int months) 
        => new($"=EOMONTH({startDate},{months})");

    public static CellFormula Weekday(string date, int? returnType) 
        => new(returnType.HasValue
        ? $"=WEEKDAY({date},{returnType.Value})"
        : $"=WEEKDAY({date})");

    public static CellFormula Round(string number, int numDigits) 
        => new($"=ROUND({number},{numDigits})");

    public static CellFormula RoundUp(string number, int numDigits) 
        => new($"=ROUNDUP({number},{numDigits})");

    public static CellFormula RoundDown(string number, int numDigits) 
        => new($"=ROUNDDOWN({number},{numDigits})");

    public static CellFormula Abs(string number) 
        => new($"=ABS({number})");

    public static CellFormula Power(string number, string power) 
        => new($"=POWER({number},{power})");

    public static CellFormula Sqrt(string number) 
        => new($"=SQRT({number})");

    public static CellFormula Mod(string number, string divisor) 
        => new($"=MOD({number},{divisor})");

    public static CellFormula Ceiling(string number, string? significance) 
        => new(significance != null
        ? $"=CEILING({number},{significance})"
        : $"=CEILING({number})");

    public static CellFormula Floor(string number, string? significance) 
        => new(significance != null
        ? $"=FLOOR({number},{significance})"
        : $"=FLOOR({number})");

    public static CellFormula Rand() => new("=RAND()");

    public static CellFormula RandBetween(int bottom, int top) 
        => new($"=RANDBETWEEN({bottom},{top})");
}
