namespace FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

public static class FormulaBuilder
{
    public static CellFormula VLookup(string lookupValue, string tableRange, int columnIndex, bool exactMatch = false)
    {
        var matchType = exactMatch ? "FALSE" : "TRUE";
        return new($"=VLOOKUP({lookupValue},{tableRange},{columnIndex},{matchType})");
    }

    public static CellFormula XLookup(string lookupValue, string lookupArray, string returnArray, string? ifNotFound = null)
    {
        var formula = ifNotFound != null
            ? $"=XLOOKUP({lookupValue},{lookupArray},{returnArray},{ifNotFound})"
            : $"=XLOOKUP({lookupValue},{lookupArray},{returnArray})";
        return new(formula);
    }

    public static CellFormula Index(string array, int rowNum, int? columnNum = null)
    {
        var formula = columnNum.HasValue
            ? $"=INDEX({array},{rowNum},{columnNum.Value})"
            : $"=INDEX({array},{rowNum})";
        return new(formula);
    }

    public static CellFormula Match(string lookupValue, string lookupArray, int matchType = 0)
    {
        return new($"=MATCH({lookupValue},{lookupArray},{matchType})");
    }

    public static CellFormula IndexMatch(string returnArray, string lookupValue, string lookupArray, int matchType = 0)
    {
        return new($"=INDEX({returnArray},MATCH({lookupValue},{lookupArray},{matchType}))");
    }

    public static CellFormula Concatenate(params string[] values)
    {
        return new($"=CONCATENATE({string.Join(",", values)})");
    }

    public static CellFormula Concat(params string[] values)
    {
        return new($"=CONCAT({string.Join(",", values)})");
    }

    public static CellFormula Left(string text, int numChars)
    {
        return new($"=LEFT({text},{numChars})");
    }

    public static CellFormula Right(string text, int numChars)
    {
        return new($"=RIGHT({text},{numChars})");
    }

    public static CellFormula Mid(string text, int startNum, int numChars)
    {
        return new($"=MID({text},{startNum},{numChars})");
    }

    public static CellFormula Trim(string text)
    {
        return new($"=TRIM({text})");
    }

    public static CellFormula Upper(string text)
    {
        return new($"=UPPER({text})");
    }

    public static CellFormula Lower(string text)
    {
        return new($"=LOWER({text})");
    }

    public static CellFormula Proper(string text)
    {
        return new($"=PROPER({text})");
    }

    public static CellFormula Len(string text)
    {
        return new($"=LEN({text})");
    }

    public static CellFormula If(string condition, string valueIfTrue, string valueIfFalse)
    {
        return new($"=IF({condition},{valueIfTrue},{valueIfFalse})");
    }

    public static CellFormula And(params string[] conditions)
    {
        return new($"=AND({string.Join(",", conditions)})");
    }

    public static CellFormula Or(params string[] conditions)
    {
        return new($"=OR({string.Join(",", conditions)})");
    }

    public static CellFormula Not(string condition)
    {
        return new($"=NOT({condition})");
    }

    public static CellFormula Xor(params string[] conditions)
    {
        return new($"=XOR({string.Join(",", conditions)})");
    }

    public static CellFormula Today()
    {
        return new("=TODAY()");
    }

    public static CellFormula Now()
    {
        return new("=NOW()");
    }

    public static CellFormula Date(int year, int month, int day)
    {
        return new($"=DATE({year},{month},{day})");
    }

    public static CellFormula Year(string date)
    {
        return new($"=YEAR({date})");
    }

    public static CellFormula Month(string date)
    {
        return new($"=MONTH({date})");
    }

    public static CellFormula Day(string date)
    {
        return new($"=DAY({date})");
    }

    public static CellFormula DateDif(string startDate, string endDate, string unit)
    {
        return new($"=DATEDIF({startDate},{endDate},\"{unit}\")");
    }

    public static CellFormula EoMonth(string startDate, int months)
    {
        return new($"=EOMONTH({startDate},{months})");
    }

    public static CellFormula Weekday(string date, int? returnType = null)
    {
        var formula = returnType.HasValue
            ? $"=WEEKDAY({date},{returnType.Value})"
            : $"=WEEKDAY({date})";
        return new(formula);
    }

    public static CellFormula Round(string number, int numDigits)
    {
        return new($"=ROUND({number},{numDigits})");
    }

    public static CellFormula RoundUp(string number, int numDigits)
    {
        return new($"=ROUNDUP({number},{numDigits})");
    }

    public static CellFormula RoundDown(string number, int numDigits)
    {
        return new($"=ROUNDDOWN({number},{numDigits})");
    }

    public static CellFormula Abs(string number)
    {
        return new($"=ABS({number})");
    }

    public static CellFormula Power(string number, string power)
    {
        return new($"=POWER({number},{power})");
    }

    public static CellFormula Sqrt(string number)
    {
        return new($"=SQRT({number})");
    }

    public static CellFormula Mod(string number, string divisor)
    {
        return new($"=MOD({number},{divisor})");
    }

    public static CellFormula Ceiling(string number, string? significance = null)
    {
        var formula = significance != null
            ? $"=CEILING({number},{significance})"
            : $"=CEILING({number})";
        return new(formula);
    }

    public static CellFormula Floor(string number, string? significance = null)
    {
        var formula = significance != null
            ? $"=FLOOR({number},{significance})"
            : $"=FLOOR({number})";
        return new(formula);
    }

    public static CellFormula Rand()
    {
        return new("=RAND()");
    }

    public static CellFormula RandBetween(int bottom, int top)
    {
        return new($"=RANDBETWEEN({bottom},{top})");
    }
}
