using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class WorkBookEdgeCasesTests
{
    [Fact]
    public void SaveToFile_WithInvalidPath_ThrowsException()
    {
        var workbook = new WorkBook("Test", [new WorkSheet("Sheet1")]);
        
        Assert.Throws<DirectoryNotFoundException>(() => 
            workbook.SaveToFile("/invalid/nonexistent/path/file.xlsx"));
    }

    [Fact]
    public void SaveToFile_WithValidPath_CreatesFile()
    {
        var sheet = new WorkSheet("Sheet1");
        sheet.AddCell(new(0, 0), "Test", null);
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            
            Assert.True(File.Exists(tempPath));
        }
        finally
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }

    [Fact]
    public void Constructor_WithEmptySheets_CreatesWorkbook()
    {
        var workbook = new WorkBook("Empty", []);
        
        Assert.Equal("Empty", workbook.Name);
        Assert.Empty(workbook.Sheets);
    }

    [Fact]
    public void Constructor_WithEmptyName_CreatesWorkbook()
    {
        var workbook = new WorkBook("", [new WorkSheet("Sheet1")]);
        
        Assert.Equal("", workbook.Name);
    }

    [Fact]
    public void Constructor_WithWhitespaceName_CreatesWorkbook()
    {
        var workbook = new WorkBook("   ", [new WorkSheet("Sheet1")]);
        
        Assert.Equal("   ", workbook.Name);
    }

    [Fact]
    public void Constructor_WithSingleSheet_StoresSheet()
    {
        var sheet = new WorkSheet("Sheet1");
        var workbook = new WorkBook("Test", [sheet]);
        
        Assert.Single(workbook.Sheets);
        Assert.Equal("Sheet1", workbook.Sheets.First().Name);
    }

    [Fact]
    public void Constructor_WithMultipleSheets_StoresAllSheets()
    {
        var sheet1 = new WorkSheet("Sheet1");
        var sheet2 = new WorkSheet("Sheet2");
        var sheet3 = new WorkSheet("Sheet3");
        var workbook = new WorkBook("Test", [sheet1, sheet2, sheet3]);
        
        Assert.Equal(3, workbook.Sheets.Count());
    }

    [Fact]
    public void Constructor_WithDuplicateSheetNames_AllowsDuplicates()
    {
        var sheet1 = new WorkSheet("Sheet");
        var sheet2 = new WorkSheet("Sheet");
        var workbook = new WorkBook("Test", [sheet1, sheet2]);
        
        Assert.Equal(2, workbook.Sheets.Count());
    }
}
