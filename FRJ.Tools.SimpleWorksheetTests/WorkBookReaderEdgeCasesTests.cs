using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class WorkBookReaderEdgeCasesTests
{
    [Fact]
    public void LoadFromFile_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        const string nonExistentPath = "/tmp/nonexistent_workbook_12345.xlsx";
        
        Assert.Throws<FileNotFoundException>(() => 
            WorkBookReader.LoadFromFile(nonExistentPath));
    }

    [Fact]
    public void LoadFromFile_WithInvalidPath_ThrowsException()
    {
        const string invalidPath = "";
        
        Assert.Throws<ArgumentException>(() => 
            WorkBookReader.LoadFromFile(invalidPath));
    }

    [Fact]
    public void LoadFromFile_WithNonExcelFile_ThrowsException()
    {
        var tempFile = Path.GetTempFileName();
        
        try
        {
            File.WriteAllText(tempFile, "not an excel file");
            
            Assert.Throws<FileFormatException>(() => 
                WorkBookReader.LoadFromFile(tempFile));
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Fact]
    public void LoadFromStream_WithClosedStream_ThrowsException()
    {
        var stream = new MemoryStream();
        stream.Close();
        
        Assert.Throws<DocumentFormat.OpenXml.Packaging.OpenXmlPackageException>(() => 
            WorkBookReader.LoadFromStream(stream));
    }

    [Fact]
    public void LoadFromStream_WithEmptyStream_ThrowsException()
    {
        using var stream = new MemoryStream();
        
        Assert.Throws<FileFormatException>(() => 
            WorkBookReader.LoadFromStream(stream));
    }

    [Fact]
    public void LoadFromStream_WithCorruptData_ThrowsException()
    {
        using var stream = new MemoryStream([0x00, 0x01, 0x02, 0x03, 0x04]);
        
        Assert.Throws<FileFormatException>(() => 
            WorkBookReader.LoadFromStream(stream));
    }

    [Fact]
    public void LoadFromFile_WithMultipleSheets_LoadsAllSheets()
    {
        var sheet1 = new WorkSheet("Sheet1");
        var sheet2 = new WorkSheet("Sheet2");
        var sheet3 = new WorkSheet("Sheet3");
        
        sheet1.AddCell(new(0, 0), "Data1", null);
        sheet2.AddCell(new(0, 0), "Data2", null);
        sheet3.AddCell(new(0, 0), "Data3", null);
        
        WorkBook workbook = new("Multi", [sheet1, sheet2, sheet3]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var loaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.Equal(3, loaded.Sheets.Count());
            Assert.Equal("Sheet1", loaded.Sheets.ElementAt(0).Name);
            Assert.Equal("Sheet2", loaded.Sheets.ElementAt(1).Name);
            Assert.Equal("Sheet3", loaded.Sheets.ElementAt(2).Name);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LoadFromFile_WithEmptySheet_HandlesCorrectly()
    {
        var emptySheet = new WorkSheet("Empty");
        WorkBook workbook = new("Test", [emptySheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var loaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.Single(loaded.Sheets);
            Assert.Equal("Empty", loaded.Sheets.First().Name);
            Assert.Empty(loaded.Sheets.First().Cells.Cells);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LoadFromFile_WithVeryLongSheetName_HandlesCorrectly()
    {
        var longName = new string('A', 31);
        var sheet = new WorkSheet(longName);
        sheet.AddCell(new(0, 0), "Test", null);
        
        WorkBook workbook = new("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var loaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.Equal(longName, loaded.Sheets.First().Name);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LoadFromFile_WithSpecialCharactersInCellValue_PreservesCharacters()
    {
        var sheet = new WorkSheet("Special");
        sheet.AddCell(new(0, 0), "Hello 世界 🌍 !@#$%", null);
        
        WorkBook workbook = new("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var loaded = WorkBookReader.LoadFromFile(tempPath);
            
            var cell = loaded.Sheets.First().Cells.Cells[new(0, 0)];
            Assert.True(cell.Value.Value.IsT2);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }
}
