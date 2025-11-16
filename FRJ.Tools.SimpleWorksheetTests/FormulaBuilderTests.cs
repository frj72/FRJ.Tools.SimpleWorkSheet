using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class FormulaBuilderTests
{
    [Fact]
    public void VLookup_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.VLookup("A2", "D1:E10", 2, true);
        
        Assert.Equal("=VLOOKUP(A2,D1:E10,2,FALSE)", formula.Value);
    }

    [Fact]
    public void VLookup_ApproximateMatch_UsesTrue()
    {
        var formula = FormulaBuilder.VLookup("A2", "D1:E10", 2);
        
        Assert.Equal("=VLOOKUP(A2,D1:E10,2,TRUE)", formula.Value);
    }

    [Fact]
    public void XLookup_WithoutIfNotFound_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.XLookup("A2", "B1:B10", "C1:C10");
        
        Assert.Equal("=XLOOKUP(A2,B1:B10,C1:C10)", formula.Value);
    }

    [Fact]
    public void XLookup_WithIfNotFound_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.XLookup("A2", "B1:B10", "C1:C10", "\"Not Found\"");
        
        Assert.Equal("=XLOOKUP(A2,B1:B10,C1:C10,\"Not Found\")", formula.Value);
    }

    [Fact]
    public void Index_WithRowOnly_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Index("A1:C10", 5);
        
        Assert.Equal("=INDEX(A1:C10,5)", formula.Value);
    }

    [Fact]
    public void Index_WithRowAndColumn_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Index("A1:C10", 5, 2);
        
        Assert.Equal("=INDEX(A1:C10,5,2)", formula.Value);
    }

    [Fact]
    public void Match_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Match("A2", "B1:B10", 0);
        
        Assert.Equal("=MATCH(A2,B1:B10,0)", formula.Value);
    }

    [Fact]
    public void IndexMatch_CombinesFormulas()
    {
        var formula = FormulaBuilder.IndexMatch("C1:C10", "A2", "B1:B10", 0);
        
        Assert.Equal("=INDEX(C1:C10,MATCH(A2,B1:B10,0))", formula.Value);
    }

    [Fact]
    public void Concatenate_JoinsValues()
    {
        var formula = FormulaBuilder.Concatenate("A1", "\" \"", "B1");
        
        Assert.Equal("=CONCATENATE(A1,\" \",B1)", formula.Value);
    }

    [Fact]
    public void Concat_JoinsValues()
    {
        var formula = FormulaBuilder.Concat("A1", "B1", "C1");
        
        Assert.Equal("=CONCAT(A1,B1,C1)", formula.Value);
    }

    [Fact]
    public void Left_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Left("A1", 5);
        
        Assert.Equal("=LEFT(A1,5)", formula.Value);
    }

    [Fact]
    public void Right_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Right("A1", 3);
        
        Assert.Equal("=RIGHT(A1,3)", formula.Value);
    }

    [Fact]
    public void Mid_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Mid("A1", 2, 5);
        
        Assert.Equal("=MID(A1,2,5)", formula.Value);
    }

    [Fact]
    public void Trim_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Trim("A1");
        
        Assert.Equal("=TRIM(A1)", formula.Value);
    }

    [Fact]
    public void Upper_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Upper("A1");
        
        Assert.Equal("=UPPER(A1)", formula.Value);
    }

    [Fact]
    public void Lower_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Lower("A1");
        
        Assert.Equal("=LOWER(A1)", formula.Value);
    }

    [Fact]
    public void Proper_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Proper("A1");
        
        Assert.Equal("=PROPER(A1)", formula.Value);
    }

    [Fact]
    public void Len_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Len("A1");
        
        Assert.Equal("=LEN(A1)", formula.Value);
    }

    [Fact]
    public void If_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.If("A1>100", "\"High\"", "\"Low\"");
        
        Assert.Equal("=IF(A1>100,\"High\",\"Low\")", formula.Value);
    }

    [Fact]
    public void And_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.And("A1>10", "B1<20", "C1=5");
        
        Assert.Equal("=AND(A1>10,B1<20,C1=5)", formula.Value);
    }

    [Fact]
    public void Or_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Or("A1>10", "B1<20");
        
        Assert.Equal("=OR(A1>10,B1<20)", formula.Value);
    }

    [Fact]
    public void Not_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Not("A1>10");
        
        Assert.Equal("=NOT(A1>10)", formula.Value);
    }

    [Fact]
    public void Xor_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Xor("A1>10", "B1<20");
        
        Assert.Equal("=XOR(A1>10,B1<20)", formula.Value);
    }

    [Fact]
    public void Today_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Today();
        
        Assert.Equal("=TODAY()", formula.Value);
    }

    [Fact]
    public void Now_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Now();
        
        Assert.Equal("=NOW()", formula.Value);
    }

    [Fact]
    public void Date_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Date(2025, 11, 16);
        
        Assert.Equal("=DATE(2025,11,16)", formula.Value);
    }

    [Fact]
    public void Year_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Year("A1");
        
        Assert.Equal("=YEAR(A1)", formula.Value);
    }

    [Fact]
    public void Month_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Month("A1");
        
        Assert.Equal("=MONTH(A1)", formula.Value);
    }

    [Fact]
    public void Day_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Day("A1");
        
        Assert.Equal("=DAY(A1)", formula.Value);
    }

    [Fact]
    public void DateDif_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.DateDif("A1", "B1", "Y");
        
        Assert.Equal("=DATEDIF(A1,B1,\"Y\")", formula.Value);
    }

    [Fact]
    public void EoMonth_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.EoMonth("A1", 3);
        
        Assert.Equal("=EOMONTH(A1,3)", formula.Value);
    }

    [Fact]
    public void Weekday_WithoutReturnType_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Weekday("A1");
        
        Assert.Equal("=WEEKDAY(A1)", formula.Value);
    }

    [Fact]
    public void Weekday_WithReturnType_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Weekday("A1", 2);
        
        Assert.Equal("=WEEKDAY(A1,2)", formula.Value);
    }

    [Fact]
    public void Round_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Round("A1", 2);
        
        Assert.Equal("=ROUND(A1,2)", formula.Value);
    }

    [Fact]
    public void RoundUp_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.RoundUp("A1", 2);
        
        Assert.Equal("=ROUNDUP(A1,2)", formula.Value);
    }

    [Fact]
    public void RoundDown_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.RoundDown("A1", 2);
        
        Assert.Equal("=ROUNDDOWN(A1,2)", formula.Value);
    }

    [Fact]
    public void Abs_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Abs("A1");
        
        Assert.Equal("=ABS(A1)", formula.Value);
    }

    [Fact]
    public void Power_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Power("A1", "2");
        
        Assert.Equal("=POWER(A1,2)", formula.Value);
    }

    [Fact]
    public void Sqrt_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Sqrt("A1");
        
        Assert.Equal("=SQRT(A1)", formula.Value);
    }

    [Fact]
    public void Mod_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Mod("A1", "3");
        
        Assert.Equal("=MOD(A1,3)", formula.Value);
    }

    [Fact]
    public void Ceiling_WithoutSignificance_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Ceiling("A1");
        
        Assert.Equal("=CEILING(A1)", formula.Value);
    }

    [Fact]
    public void Ceiling_WithSignificance_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Ceiling("A1", "0.5");
        
        Assert.Equal("=CEILING(A1,0.5)", formula.Value);
    }

    [Fact]
    public void Floor_WithoutSignificance_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Floor("A1");
        
        Assert.Equal("=FLOOR(A1)", formula.Value);
    }

    [Fact]
    public void Floor_WithSignificance_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Floor("A1", "0.5");
        
        Assert.Equal("=FLOOR(A1,0.5)", formula.Value);
    }

    [Fact]
    public void Rand_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.Rand();
        
        Assert.Equal("=RAND()", formula.Value);
    }

    [Fact]
    public void RandBetween_GeneratesCorrectFormula()
    {
        var formula = FormulaBuilder.RandBetween(1, 100);
        
        Assert.Equal("=RANDBETWEEN(1,100)", formula.Value);
    }
}
