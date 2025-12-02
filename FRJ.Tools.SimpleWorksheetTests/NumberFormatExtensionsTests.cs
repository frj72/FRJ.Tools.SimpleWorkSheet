using FRJ.Tools.SimpleWorkSheet.Components.Formatting;

namespace FRJ.Tools.SimpleWorksheetTests;

public class NumberFormatExtensionsTests
{
    [Fact]
    public void ToFormatString_WithInteger_ReturnsIntegerFormat()
    {
        var format = NumberFormat.Integer;

        var result = format.ToFormatString();

        Assert.Equal("0", result);
    }

    [Fact]
    public void ToFormatString_WithFloat2_ReturnsFloat2Format()
    {
        var format = NumberFormat.Float2;

        var result = format.ToFormatString();

        Assert.Equal("F2", result);
    }

    [Fact]
    public void ToFormatString_WithFloat3_ReturnsFloat3Format()
    {
        var format = NumberFormat.Float3;

        var result = format.ToFormatString();

        Assert.Equal("F3", result);
    }

    [Fact]
    public void ToFormatString_WithFloat4_ReturnsFloat4Format()
    {
        var format = NumberFormat.Float4;

        var result = format.ToFormatString();

        Assert.Equal("F4", result);
    }

    [Fact]
    public void ToExcelFormatString_WithInteger_ReturnsExcelIntegerFormat()
    {
        var format = NumberFormat.Integer;

        var result = format.ToExcelFormatString();

        Assert.Equal("0", result);
    }

    [Fact]
    public void ToExcelFormatString_WithFloat2_ReturnsExcelFloat2Format()
    {
        var format = NumberFormat.Float2;

        var result = format.ToExcelFormatString();

        Assert.Equal("0.00", result);
    }

    [Fact]
    public void ToExcelFormatString_WithFloat3_ReturnsExcelFloat3Format()
    {
        var format = NumberFormat.Float3;

        var result = format.ToExcelFormatString();

        Assert.Equal("0.000", result);
    }

    [Fact]
    public void ToExcelFormatString_WithFloat4_ReturnsExcelFloat4Format()
    {
        var format = NumberFormat.Float4;

        var result = format.ToExcelFormatString();

        Assert.Equal("0.0000", result);
    }

    [Fact]
    public void ToExcelFormatId_WithInteger_Returns169()
    {
        var format = NumberFormat.Integer;

        var result = format.ToExcelFormatId();

        Assert.Equal(169u, result);
    }

    [Fact]
    public void ToExcelFormatId_WithFloat2_Returns170()
    {
        var format = NumberFormat.Float2;

        var result = format.ToExcelFormatId();

        Assert.Equal(170u, result);
    }

    [Fact]
    public void ToExcelFormatId_WithFloat3_Returns171()
    {
        var format = NumberFormat.Float3;

        var result = format.ToExcelFormatId();

        Assert.Equal(171u, result);
    }

    [Fact]
    public void ToExcelFormatId_WithFloat4_Returns172()
    {
        var format = NumberFormat.Float4;

        var result = format.ToExcelFormatId();

        Assert.Equal(172u, result);
    }

    [Fact]
    public void ToFormatString_WithAllEnumValues_ReturnsValidFormatStrings()
    {
        var allFormats = Enum.GetValues<NumberFormat>();

        foreach (var format in allFormats)
        {
            var result = format.ToFormatString();

            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.True(result.Contains('0') || result.Contains('F'));
        }
    }

    [Fact]
    public void ToExcelFormatString_WithAllEnumValues_ReturnsValidFormatStrings()
    {
        var allFormats = Enum.GetValues<NumberFormat>();

        foreach (var format in allFormats)
        {
            var result = format.ToExcelFormatString();

            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.Contains('0', result);
        }
    }

    [Fact]
    public void ToExcelFormatId_WithAllEnumValues_ReturnsValidIds()
    {
        var allFormats = Enum.GetValues<NumberFormat>();

        foreach (var format in allFormats)
        {
            var result = format.ToExcelFormatId();

            Assert.InRange(result, 169u, 172u);
        }
    }

    [Fact]
    public void ToFormatString_WithInteger_ContainsZeroOnly()
    {
        var format = NumberFormat.Integer;
        var result = format.ToFormatString();

        Assert.Contains('0', result);
        Assert.DoesNotContain('F', result);
        Assert.DoesNotContain('.', result);
    }

    [Fact]
    public void ToFormatString_WithFloatFormats_ContainsFAndNumber()
    {
        var float2 = NumberFormat.Float2.ToFormatString();
        var float3 = NumberFormat.Float3.ToFormatString();
        var float4 = NumberFormat.Float4.ToFormatString();

        Assert.Contains('F', float2);
        Assert.Contains('F', float3);
        Assert.Contains('F', float4);
        Assert.Contains('2', float2);
        Assert.Contains('3', float3);
        Assert.Contains('4', float4);
    }

    [Fact]
    public void ToExcelFormatString_WithInteger_ContainsZeroOnly()
    {
        var format = NumberFormat.Integer;
        var result = format.ToExcelFormatString();

        Assert.Contains('0', result);
        Assert.DoesNotContain('.', result);
    }

    [Fact]
    public void ToExcelFormatString_WithFloatFormats_ContainsDecimalPlaces()
    {
        var float2 = NumberFormat.Float2.ToExcelFormatString();
        var float3 = NumberFormat.Float3.ToExcelFormatString();
        var float4 = NumberFormat.Float4.ToExcelFormatString();

        Assert.Contains('.', float2);
        Assert.Contains('.', float3);
        Assert.Contains('.', float4);

        Assert.Equal(3, float2.Count(c => c == '0'));
        Assert.Equal(4, float3.Count(c => c == '0'));
        Assert.Equal(5, float4.Count(c => c == '0'));
    }

    [Fact]
    public void ToExcelFormatId_ReturnsSequentialIds()
    {
        var formats = new[]
        {
            NumberFormat.Integer,
            NumberFormat.Float2,
            NumberFormat.Float3,
            NumberFormat.Float4
        };

        var expectedIds = new uint[] { 169, 170, 171, 172 };

        for (var i = 0; i < formats.Length; i++)
        {
            var result = formats[i].ToExcelFormatId();
            Assert.Equal(expectedIds[i], result);
        }
    }

    [Fact]
    public void ToFormatString_ReturnsSameResultForSameFormat()
    {
        var format = NumberFormat.Float2;

        var result1 = format.ToFormatString();
        var result2 = format.ToFormatString();

        Assert.Equal(result1, result2);
    }

    [Fact]
    public void ToExcelFormatString_ReturnsSameResultForSameFormat()
    {
        var format = NumberFormat.Float2;

        var result1 = format.ToExcelFormatString();
        var result2 = format.ToExcelFormatString();

        Assert.Equal(result1, result2);
    }

    [Fact]
    public void ToExcelFormatId_ReturnsSameResultForSameFormat()
    {
        var format = NumberFormat.Float2;

        var result1 = format.ToExcelFormatId();
        var result2 = format.ToExcelFormatId();

        Assert.Equal(result1, result2);
    }

    [Fact]
    public void ToFormatString_AllFormatsAreValidDotNetFormats()
    {
        var testNumber = 123.456789;
        var allFormats = Enum.GetValues<NumberFormat>();

        foreach (var format in allFormats)
        {
            var formatString = format.ToFormatString();

            var formatted = testNumber.ToString(formatString);
            Assert.NotNull(formatted);
            Assert.NotEmpty(formatted);
        }
    }

    [Fact]
    public void ToExcelFormatId_AllIdsAreUnique()
    {
        var allFormats = Enum.GetValues<NumberFormat>();
        var ids = allFormats.Select(f => f.ToExcelFormatId()).ToList();

        var uniqueIds = ids.Distinct().ToList();

        Assert.Equal(ids.Count, uniqueIds.Count);
    }

    [Fact]
    public void ToExcelFormatId_AllIdsAreInValidRange()
    {
        var allFormats = Enum.GetValues<NumberFormat>();

        foreach (var format in allFormats)
        {
            var id = format.ToExcelFormatId();

            Assert.True(id >= 169);
            Assert.True(id <= 200);
        }
    }

    [Fact]
    public void ToFormatString_IntegerFormat_ProducesWholeNumbers()
    {
        var format = NumberFormat.Integer;
        var formatString = format.ToFormatString();

        var result = 123.789.ToString(formatString);

        Assert.Equal("124", result);
    }

    [Fact]
    public void ToFormatString_Float2Format_ProducesTwoDecimalPlaces()
    {
        var format = NumberFormat.Float2;
        var formatString = format.ToFormatString();

        var result = 123.456.ToString(formatString);

        Assert.Equal("123.46", result);
    }

    [Fact]
    public void ToFormatString_Float3Format_ProducesThreeDecimalPlaces()
    {
        var format = NumberFormat.Float3;
        var formatString = format.ToFormatString();

        var result = 123.4567.ToString(formatString);

        Assert.Equal("123.457", result);
    }

    [Fact]
    public void ToFormatString_Float4Format_ProducesFourDecimalPlaces()
    {
        var format = NumberFormat.Float4;
        var formatString = format.ToFormatString();

        var result = 123.45678.ToString(formatString);

        Assert.Equal("123.4568", result);
    }
}