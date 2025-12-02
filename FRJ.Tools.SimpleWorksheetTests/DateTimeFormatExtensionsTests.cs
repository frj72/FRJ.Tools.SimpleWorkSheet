using FRJ.Tools.SimpleWorkSheet.Components.Formatting;

namespace FRJ.Tools.SimpleWorksheetTests;

public class DateTimeFormatExtensionsTests
{
    [Fact]
    public void ToFormatString_WithIsoDateTime_ReturnsIsoDateTimeFormat()
    {
        var format = DateFormat.IsoDateTime;
        
        var result = format.ToFormatString();
        
        Assert.Equal("yyyy-MM-dd HH:mm:ss", result);
    }
    
    [Fact]
    public void ToFormatString_WithIsoDate_ReturnsIsoDateFormat()
    {
        var format = DateFormat.IsoDate;
        
        var result = format.ToFormatString();
        
        Assert.Equal("yyyy-MM-dd", result);
    }
    
    [Fact]
    public void ToFormatString_WithDateTime_ReturnsDateTimeFormat()
    {
        var format = DateFormat.DateTime;
        
        var result = format.ToFormatString();
        
        Assert.Equal("dd/MM/yyyy HH:mm:ss", result);
    }
    
    [Fact]
    public void ToFormatString_WithDateOnly_ReturnsDateOnlyFormat()
    {
        var format = DateFormat.DateOnly;
        
        var result = format.ToFormatString();
        
        Assert.Equal("dd/MM/yyyy", result);
    }
    
    [Fact]
    public void ToFormatString_WithTimeOnly_ReturnsTimeOnlyFormat()
    {
        var format = DateFormat.TimeOnly;
        
        var result = format.ToFormatString();
        
        Assert.Equal("HH:mm:ss", result);
    }
    
    [Fact]
    public void ToExcelFormatString_WithIsoDateTime_ReturnsExcelIsoDateTimeFormat()
    {
        var format = DateFormat.IsoDateTime;
        
        var result = format.ToExcelFormatString();
        
        Assert.Equal("yyyy-mm-dd hh:mm:ss", result);
    }
    
    [Fact]
    public void ToExcelFormatString_WithIsoDate_ReturnsExcelIsoDateFormat()
    {
        var format = DateFormat.IsoDate;
        
        var result = format.ToExcelFormatString();
        
        Assert.Equal("yyyy-mm-dd", result);
    }
    
    [Fact]
    public void ToExcelFormatString_WithDateTime_ReturnsExcelDateTimeFormat()
    {
        var format = DateFormat.DateTime;
        
        var result = format.ToExcelFormatString();
        
        Assert.Equal("dd/mm/yyyy hh:mm:ss", result);
    }
    
    [Fact]
    public void ToExcelFormatString_WithDateOnly_ReturnsExcelDateOnlyFormat()
    {
        var format = DateFormat.DateOnly;
        
        var result = format.ToExcelFormatString();
        
        Assert.Equal("dd/mm/yyyy", result);
    }
    
    [Fact]
    public void ToExcelFormatString_WithTimeOnly_ReturnsExcelTimeOnlyFormat()
    {
        var format = DateFormat.TimeOnly;
        
        var result = format.ToExcelFormatString();
        
        Assert.Equal("hh:mm:ss", result);
    }
    
    [Fact]
    public void ToExcelFormatId_WithIsoDateTime_Returns164()
    {
        var format = DateFormat.IsoDateTime;
        
        var result = format.ToExcelFormatId();
        
        Assert.Equal(164u, result);
    }
    
    [Fact]
    public void ToExcelFormatId_WithIsoDate_Returns165()
    {
        var format = DateFormat.IsoDate;
        
        var result = format.ToExcelFormatId();
        
        Assert.Equal(165u, result);
    }
    
    [Fact]
    public void ToExcelFormatId_WithDateTime_Returns166()
    {
        var format = DateFormat.DateTime;
        
        var result = format.ToExcelFormatId();
        
        Assert.Equal(166u, result);
    }
    
    [Fact]
    public void ToExcelFormatId_WithDateOnly_Returns167()
    {
        var format = DateFormat.DateOnly;
        
        var result = format.ToExcelFormatId();
        
        Assert.Equal(167u, result);
    }
    
    [Fact]
    public void ToExcelFormatId_WithTimeOnly_Returns168()
    {
        var format = DateFormat.TimeOnly;
        
        var result = format.ToExcelFormatId();
        
        Assert.Equal(168u, result);
    }
    
    [Fact]
    public void ToFormatString_WithAllEnumValues_ReturnsValidFormatStrings()
    {
        var allFormats = Enum.GetValues<DateFormat>();

        foreach (var format in allFormats)
        {
            var result = format.ToFormatString();

            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.True(result.Contains('-') || result.Contains('/') || result.Contains(':') || result.Contains(' '));
        }
    }
    
    [Fact]
    public void ToExcelFormatString_WithAllEnumValues_ReturnsValidFormatStrings()
    {
        var allFormats = Enum.GetValues<DateFormat>();

        foreach (var format in allFormats)
        {
            var result = format.ToExcelFormatString();

            Assert.NotNull(result);
            Assert.NotEmpty(result);
            Assert.True(result.Contains('-') || result.Contains('/') || result.Contains(':') || result.Contains(' '));
        }
    }
    
    [Fact]
    public void ToExcelFormatId_WithAllEnumValues_ReturnsValidIds()
    {
        var allFormats = Enum.GetValues<DateFormat>();
        
        foreach (var format in allFormats)
        {
            var result = format.ToExcelFormatId();

            Assert.InRange(result, 164u, 168u);
        }
    }
    
    [Fact]
    public void ToFormatString_WithIsoDateTime_ContainsAllComponents()
    {
        var format = DateFormat.IsoDateTime;
        var result = format.ToFormatString();
        
        Assert.Contains("yyyy", result);
        Assert.Contains("MM", result);
        Assert.Contains("dd", result);
        Assert.Contains("HH", result);
        Assert.Contains("mm", result);
        Assert.Contains("ss", result);
    }
    
    [Fact]
    public void ToFormatString_WithIsoDate_ContainsDateComponentsOnly()
    {
        var format = DateFormat.IsoDate;
        var result = format.ToFormatString();
        
        Assert.Contains("yyyy", result);
        Assert.Contains("MM", result);
        Assert.Contains("dd", result);
        Assert.DoesNotContain("HH", result);
        Assert.DoesNotContain("mm", result);
        Assert.DoesNotContain("ss", result);
    }
    
    [Fact]
    public void ToFormatString_WithTimeOnly_ContainsTimeComponentsOnly()
    {
        var format = DateFormat.TimeOnly;
        var result = format.ToFormatString();
        
        Assert.DoesNotContain("yyyy", result);
        Assert.DoesNotContain("MM", result);
        Assert.DoesNotContain("dd", result);
        Assert.Contains("HH", result);
        Assert.Contains("mm", result);
        Assert.Contains("ss", result);
    }
    
    [Fact]
    public void ToExcelFormatString_UsesLowercaseForExcelCompatibility()
    {
        var format = DateFormat.IsoDateTime;
        var result = format.ToExcelFormatString();
        
        Assert.Contains("yyyy", result);
        Assert.Contains("mm", result);
        Assert.Contains("dd", result);
        Assert.Contains("hh", result);
        Assert.Contains("ss", result);
    }
    
    [Fact]
    public void ToExcelFormatId_ReturnsSequentialIds()
    {
        var formats = new[]
        {
            DateFormat.IsoDateTime,
            DateFormat.IsoDate,
            DateFormat.DateTime,
            DateFormat.DateOnly,
            DateFormat.TimeOnly
        };
        
        var expectedIds = new uint[] { 164, 165, 166, 167, 168 };
        
        for (var i = 0; i < formats.Length; i++)
        {
            var result = formats[i].ToExcelFormatId();
            Assert.Equal(expectedIds[i], result);
        }
    }
    
    [Fact]
    public void ToFormatString_ReturnsSameResultForSameFormat()
    {
        var format = DateFormat.DateTime;
        
        var result1 = format.ToFormatString();
        var result2 = format.ToFormatString();
        
        Assert.Equal(result1, result2);
    }
    
    [Fact]
    public void ToExcelFormatString_ReturnsSameResultForSameFormat()
    {
        var format = DateFormat.DateTime;
        
        var result1 = format.ToExcelFormatString();
        var result2 = format.ToExcelFormatString();
        
        Assert.Equal(result1, result2);
    }
    
    [Fact]
    public void ToExcelFormatId_ReturnsSameResultForSameFormat()
    {
        var format = DateFormat.DateTime;
        
        var result1 = format.ToExcelFormatId();
        var result2 = format.ToExcelFormatId();
        
        Assert.Equal(result1, result2);
    }
    
    [Fact]
    public void ToFormatString_WithDateTimeFormat_UsesForwardSlashes()
    {
        var format = DateFormat.DateTime;
        var result = format.ToFormatString();
        
        Assert.Contains("/", result);
        Assert.DoesNotContain("-", result);
    }
    
    [Fact]
    public void ToFormatString_WithIsoFormats_UsesHyphens()
    {
        var isoDateTime = DateFormat.IsoDateTime.ToFormatString();
        var isoDate = DateFormat.IsoDate.ToFormatString();
        
        Assert.Contains("-", isoDateTime);
        Assert.Contains("-", isoDate);
        Assert.DoesNotContain("/", isoDateTime);
        Assert.DoesNotContain("/", isoDate);
    }
    
    [Fact]
    public void ToExcelFormatString_WithDateTimeFormat_UsesForwardSlashes()
    {
        var format = DateFormat.DateTime;
        var result = format.ToExcelFormatString();
        
        Assert.Contains("/", result);
        Assert.DoesNotContain("-", result);
    }
    
    [Fact]
    public void ToExcelFormatString_WithIsoFormats_UsesHyphens()
    {
        var isoDateTime = DateFormat.IsoDateTime.ToExcelFormatString();
        var isoDate = DateFormat.IsoDate.ToExcelFormatString();
        
        Assert.Contains("-", isoDateTime);
        Assert.Contains("-", isoDate);
        Assert.DoesNotContain("/", isoDateTime);
        Assert.DoesNotContain("/", isoDate);
    }
    
    [Fact]
    public void ToFormatString_AllFormatsAreValidDotNetFormats()
    {
        var testDate = new DateTime(2024, 12, 2, 15, 30, 45);
        var allFormats = Enum.GetValues<DateFormat>();
        
        foreach (var format in allFormats)
        {
            var formatString = format.ToFormatString();

            var formatted = testDate.ToString(formatString);
            Assert.NotNull(formatted);
            Assert.NotEmpty(formatted);
        }
    }
    
    [Fact]
    public void ToExcelFormatId_AllIdsAreUnique()
    {
        var allFormats = Enum.GetValues<DateFormat>();
        var ids = allFormats.Select(f => f.ToExcelFormatId()).ToList();
        
        var uniqueIds = ids.Distinct().ToList();
        
        Assert.Equal(ids.Count, uniqueIds.Count);
    }
    
    [Fact]
    public void ToExcelFormatId_AllIdsAreInValidRange()
    {
        var allFormats = Enum.GetValues<DateFormat>();
        
        foreach (var format in allFormats)
        {
            var id = format.ToExcelFormatId();

            Assert.True(id >= 164);
            Assert.True(id <= 200);
        }
    }
}