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
        
        sheet.AddCell(new(0, 1), "Text Value", null);
        sheet.AddCell(new(1, 1), 123.45, null);
        
        sheet.AddCell(new(0, 2), "Link", cell => cell
            .WithHyperlink("https://example.com", "Test Link"));
        
        sheet.SetColumnWidth(0, 20.0);
        sheet.SetRowHeight(0, 25.0);
        sheet.FreezePanes(1, 0);
        
        WorkBook workbook = new("TestWorkbook", [sheet]);
        return SaveWorkbook(workbook);
    }

    private static string SaveWorkbook(WorkBook workbook)
    {
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
            
            var cell00 = sheet.Cells.Cells[new(0, 0)];
            Assert.True(cell00.Value.IsString());
            
            var cell11 = sheet.Cells.Cells[new(1, 1)];
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
            
            var cell00 = sheet.Cells.Cells[new(0, 0)];
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
            
            var cell02 = sheet.Cells.Cells[new(0, 2)];
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
        originalSheet.AddCell(new(1, 0), 42, null);
        
        WorkBook originalWorkbook = new("Test", [originalSheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            originalWorkbook.SaveToFile(tempPath);
            var loadedWorkbook = WorkBookReader.LoadFromFile(tempPath);
            
            var loadedSheet = loadedWorkbook.Sheets.First();
            var cell00 = loadedSheet.Cells.Cells[new(0, 0)];
            
            Assert.True(cell00.Value.IsString());
            Assert.NotNull(cell00.Font);
            Assert.True(cell00.Font.Bold);
            Assert.Equal(14, cell00.Font.Size);
            Assert.Equal("FF0000", cell00.Color);
            
            var cell10 = loadedSheet.Cells.Cells[new(1, 0)];
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
        
        WorkBook workbook = new("Test", [sheet]);
        var tempPath = SaveWorkbook(workbook);
        
        try
        {
            var loadedWorkbook = WorkBookReader.LoadFromFile(tempPath);
            
            var loadedSheet = loadedWorkbook.Sheets.First();
            var cell = loadedSheet.Cells.Cells[new(0, 0)];
            
            Assert.Equal(HorizontalAlignment.Center, cell.Style.HorizontalAlignment);
            Assert.Equal(VerticalAlignment.Middle, cell.Style.VerticalAlignment);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void RoundTrip_PreservesAllCellBorderStyles()
    {
        var sheet = new WorkSheet("Borders");
        var styles = Enum.GetValues<CellBorderStyle>();
        uint row = 0;
        foreach (var style in styles)
        {
            sheet.AddCell(new(0, row), style.ToString(), cell => cell
                .WithBorders(CellBorders.Create(
                    CellBorder.Create("000000", style),
                    CellBorder.Create("000000", style),
                    CellBorder.Create("000000", style),
                    CellBorder.Create("000000", style))));
            row++;
        }

        var workbook = new WorkBook("Borders", [sheet]);
        var tempPath = SaveWorkbook(workbook);

        try
        {
            var loaded = WorkBookReader.LoadFromFile(tempPath);
            var loadedSheet = loaded.Sheets.First();

            row = 0;
            foreach (var style in styles)
            {
                var cell = loadedSheet.Cells.Cells[new(0, row)];
                Assert.Equal(style, cell.Borders?.Top?.Style);
                Assert.Equal(style, cell.Borders?.Left?.Style);
                row++;
            }
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void RoundTrip_PreservesHorizontalAlignmentValues()
    {
        var sheet = new WorkSheet("HorizontalAlignment");
        uint row = 0;
        foreach (var alignment in Enum.GetValues<HorizontalAlignment>())
        {
            sheet.AddCell(new(0, row), alignment.ToString(), cell => cell
                .WithStyle(style => style.WithHorizontalAlignment(alignment)));
            row++;
        }
 
        var path = SaveWorkbook(new("Alignments", [sheet]));
 
        try
        {
            var loaded = WorkBookReader.LoadFromFile(path);

            var loadedSheet = loaded.Sheets.First();
            row = 0;
            foreach (var alignment in Enum.GetValues<HorizontalAlignment>())
            {
                var cell = loadedSheet.Cells.Cells[new(0, row)];
                Assert.Equal(alignment, cell.Style.HorizontalAlignment);
                row++;
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RoundTrip_PreservesVerticalAlignmentValues()
    {
        var sheet = new WorkSheet("VerticalAlignment");
        uint row = 0;
        foreach (var alignment in Enum.GetValues<VerticalAlignment>())
        {
            sheet.AddCell(new(0, row), alignment.ToString(), cell => cell
                .WithStyle(style => style.WithVerticalAlignment(alignment)));
            row++;
        }
 
        var path = SaveWorkbook(new("Alignments", [sheet]));
 
        try
        {
            var loaded = WorkBookReader.LoadFromFile(path);

            var loadedSheet = loaded.Sheets.First();
            row = 0;
            foreach (var alignment in Enum.GetValues<VerticalAlignment>())
            {
                var cell = loadedSheet.Cells.Cells[new(0, row)];
                Assert.Equal(alignment, cell.Style.VerticalAlignment);
                row++;
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RoundTrip_PreservesTextRotationAndWrapText()
    {
        var sheet = new WorkSheet("Formatting");
        sheet.AddCell(new(0, 0), "Rotated", cell => cell
            .WithStyle(style => style.WithTextRotation(45).WithWrapText(true)));

        var path = SaveWorkbook(new("Formatting", [sheet]));
 
        try
        {
            var loaded = WorkBookReader.LoadFromFile(path);

            var cell = loaded.Sheets.First().Cells.Cells[new(0, 0)];

            Assert.Equal(45, cell.Style.TextRotation);
            Assert.True(cell.Style.WrapText);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RoundTrip_PreservesHyperlinks()
    {
        var sheet = new WorkSheet("Links");
        sheet.AddCell(new(0, 0), "Docs", cell => cell.WithHyperlink("https://docs.example.com", "Docs"));
        sheet.AddCell(new(0, 1), "Email", cell => cell.WithHyperlink("mailto:user@example.com?subject=Hello", null));

        var path = SaveWorkbook(new("Links", [sheet]));
 
        try
        {
            var loaded = WorkBookReader.LoadFromFile(path);

            var loadedSheet = loaded.Sheets.First();

            var cell0 = loadedSheet.Cells.Cells[new(0, 0)];
            var cell1 = loadedSheet.Cells.Cells[new(0, 1)];

            Assert.Equal("https://docs.example.com/", cell0.Hyperlink?.Url);
            Assert.Equal("mailto:user@example.com?subject=Hello", cell1.Hyperlink?.Url);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
