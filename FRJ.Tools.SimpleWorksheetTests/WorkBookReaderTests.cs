using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class WorkBookReaderTests
{
    private static string CreateTestExcelFile()
    {
        var sheet = new WorkSheet("TestSheet");
        
        sheet.AddCell(new(0, 0), "Header 1", cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        sheet.AddCell(new(1, 0), "Header 2", cell => cell
            .WithFont(font => font.Bold())
            .WithColor("4472C4"));
        
        sheet.AddCell(new(0, 1), "Text Value");
        sheet.AddCell(new(1, 1), 123.45);
        
        sheet.AddCell(new(0, 2), "Link", cell => cell
            .WithHyperlink("https://example.com", "Test Link"));
        
        sheet.SetColumnWith(0, 20.0);
        sheet.SetRowHeight(0, 25.0);
        sheet.FreezePanes(1, 0);
        
        var workbook = new WorkBook("TestWorkbook", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        workbook.SaveToFile(tempPath);
        return tempPath;
    }

    [Fact]
    public void LoadFromFile_ValidFile_LoadsWorkbook()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(filePath);
            
            Assert.NotNull(loadedWorkbook);
            Assert.Single(loadedWorkbook.Sheets);
            Assert.Equal("TestSheet", loadedWorkbook.Sheets.First().Name);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void LoadFromFile_ParsesCellValues()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(filePath);
            var sheet = loadedWorkbook.Sheets.First();
            
            var cell00 = sheet.Cells.Cells[new CellPosition(0, 0)];
            Assert.True(cell00.Value.IsString());
            
            var cell11 = sheet.Cells.Cells[new CellPosition(1, 1)];
            Assert.True(cell11.Value.IsLong() || cell11.Value.IsDecimal());
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void LoadFromFile_ParsesStyles()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(filePath);
            var sheet = loadedWorkbook.Sheets.First();
            
            var cell00 = sheet.Cells.Cells[new CellPosition(0, 0)];
            Assert.NotNull(cell00.Font);
            Assert.True(cell00.Font.Bold);
            Assert.Equal("4472C4", cell00.Color);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void LoadFromFile_ParsesHyperlinks()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(filePath);
            var sheet = loadedWorkbook.Sheets.First();
            
            var cell02 = sheet.Cells.Cells[new CellPosition(0, 2)];
            Assert.NotNull(cell02.Hyperlink);
            Assert.Equal("https://example.com/", cell02.Hyperlink.Url);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void LoadFromFile_ParsesColumnWidths()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(filePath);
            var sheet = loadedWorkbook.Sheets.First();
            
            Assert.True(sheet.ExplicitColumnWidths.ContainsKey(0));
            Assert.True(sheet.ExplicitColumnWidths[0].IsT0);
            Assert.Equal(20.0, sheet.ExplicitColumnWidths[0].AsT0);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void LoadFromFile_ParsesRowHeights()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(filePath);
            var sheet = loadedWorkbook.Sheets.First();
            
            Assert.True(sheet.ExplicitRowHeights.ContainsKey(0));
            Assert.True(sheet.ExplicitRowHeights[0].IsT0);
            Assert.Equal(25.0, sheet.ExplicitRowHeights[0].AsT0);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void LoadFromFile_ParsesFrozenPanes()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(filePath);
            var sheet = loadedWorkbook.Sheets.First();
            
            Assert.NotNull(sheet.FrozenPane);
            Assert.Equal(1u, sheet.FrozenPane.Row);
            Assert.Equal(0u, sheet.FrozenPane.Column);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void LoadFromStream_ValidStream_LoadsWorkbook()
    {
        var filePath = CreateTestExcelFile();
        
        try
        {
            using var stream = File.OpenRead(filePath);
            var loadedWorkbook = WorkBookReader.LoadFromStream(stream);
            
            Assert.NotNull(loadedWorkbook);
            Assert.Single(loadedWorkbook.Sheets);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void RoundTrip_PreservesData()
    {
        var originalSheet = new WorkSheet("RoundTrip");
        originalSheet.AddCell(new(0, 0), "Test", cell => cell
            .WithFont(font => font.Bold().WithSize(14))
            .WithColor("FF0000"));
        originalSheet.AddCell(new(1, 0), 42);
        
        var originalWorkbook = new WorkBook("Test", [originalSheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            originalWorkbook.SaveToFile(tempPath);
            var loadedWorkbook = WorkBookReader.LoadFromFile(tempPath);
            
            var loadedSheet = loadedWorkbook.Sheets.First();
            var cell00 = loadedSheet.Cells.Cells[new CellPosition(0, 0)];
            
            Assert.True(cell00.Value.IsString());
            Assert.NotNull(cell00.Font);
            Assert.True(cell00.Font.Bold);
            Assert.Equal(14, cell00.Font.Size);
            Assert.Equal("FF0000", cell00.Color);
            
            var cell10 = loadedSheet.Cells.Cells[new CellPosition(1, 0)];
            Assert.True(cell10.Value.IsLong() || cell10.Value.IsDecimal());
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void LoadFromFile_WithAlignment_ParsesAlignment()
    {
        var sheet = new WorkSheet("Alignment");
        sheet.AddCell(new(0, 0), "Centered", cell => cell
            .WithStyle(style => style
                .WithHorizontalAlignment(HorizontalAlignment.Center)
                .WithVerticalAlignment(VerticalAlignment.Middle)));
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var loadedWorkbook = WorkBookReader.LoadFromFile(tempPath);
            
            var loadedSheet = loadedWorkbook.Sheets.First();
            var cell = loadedSheet.Cells.Cells[new CellPosition(0, 0)];
            
            Assert.Equal(HorizontalAlignment.Center, cell.Style.HorizontalAlignment);
            Assert.Equal(VerticalAlignment.Middle, cell.Style.VerticalAlignment);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }
}
