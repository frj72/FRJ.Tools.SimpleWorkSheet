using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class OpenXmlValidationTests
{
    private static void ValidateExcelFile(string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        
        if (errors.Any())
        {
            var errorMessages = string.Join("\n", errors.Select(e => 
                $"  - {e.Description}\n    Path: {e.Path?.XPath}\n    Part: {e.Part?.Uri}"));
            Assert.Fail($"OpenXML validation failed:\n{errorMessages}");
        }
    }

    [Fact]
    public void SimplestSheet_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Hello");
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook = new WorkBook("Test", [sheet]);
            workbook.SaveToFile(tempPath);
            
            ValidateExcelFile(tempPath);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void RoundTrip_SimplestCase_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Hello");
        sheet.AddCell(new(1, 0), 123);
        
        var tempPath1 = Path.GetTempFileName() + ".xlsx";
        var tempPath2 = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook1 = new WorkBook("Test", [sheet]);
            workbook1.SaveToFile(tempPath1);
            
            ValidateExcelFile(tempPath1);
            
            var loaded = WorkBookReader.LoadFromFile(tempPath1);
            var loadedSheet = loaded.Sheets.First();
            
            var workbook2 = new WorkBook("Test", [loadedSheet]);
            workbook2.SaveToFile(tempPath2);
            
            ValidateExcelFile(tempPath2);
        }
        finally
        {
            File.Delete(tempPath1);
            File.Delete(tempPath2);
        }
    }

    [Fact]
    public void RoundTrip_WithHyperlinks_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Link", cell => cell
            .WithHyperlink("https://example.com"));
        
        var tempPath1 = Path.GetTempFileName() + ".xlsx";
        var tempPath2 = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook1 = new WorkBook("Test", [sheet]);
            workbook1.SaveToFile(tempPath1);
            
            ValidateExcelFile(tempPath1);
            
            var loaded = WorkBookReader.LoadFromFile(tempPath1);
            var loadedSheet = loaded.Sheets.First();
            
            var workbook2 = new WorkBook("Test", [loadedSheet]);
            workbook2.SaveToFile(tempPath2);
            
            ValidateExcelFile(tempPath2);
        }
        finally
        {
            File.Delete(tempPath1);
            File.Delete(tempPath2);
        }
    }

    [Fact]
    public void RoundTrip_WithAllFeatures_ValidatesCorrectly()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.AddCell(new(0, 0), "Header", cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4")
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Center)));
        
        sheet.AddCell(new(0, 1), "Link", cell => cell
            .WithHyperlink("https://example.com"));
        
        sheet.SetColumnWith(0, 20.0);
        sheet.SetRowHeight(0, 25.0);
        sheet.FreezePanes(1, 0);
        
        var tempPath1 = Path.GetTempFileName() + ".xlsx";
        var tempPath2 = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            var workbook1 = new WorkBook("Test", [sheet]);
            workbook1.SaveToFile(tempPath1);
            
            ValidateExcelFile(tempPath1);
            
            var loaded = WorkBookReader.LoadFromFile(tempPath1);
            var loadedSheet = loaded.Sheets.First();
            
            var workbook2 = new WorkBook("Test", [loadedSheet]);
            workbook2.SaveToFile(tempPath2);
            
            ValidateExcelFile(tempPath2);
        }
        finally
        {
            File.Delete(tempPath1);
            File.Delete(tempPath2);
        }
    }
}
