using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ExampleFileValidationTests
{
    private static readonly string ExamplesPath = Path.Combine(
        Directory.GetCurrentDirectory(), 
        "../../../..",
        "FRJ.Tools.SimpleWorkSheet.Examples/Output");

    [Theory]
    [InlineData("31_RowHeight.xlsx")]
    [InlineData("32_FreezePanes.xlsx")]
    [InlineData("33_TextAlignment.xlsx")]
    [InlineData("34_Hyperlinks.xlsx")]
    [InlineData("35_ReadExcel.xlsx")]
    [InlineData("36_RoundTripEditing.xlsx")]
    public void ExampleFile_CanBeLoadedByWorkBookReader(string fileName)
    {
        var filePath = Path.Combine(ExamplesPath, fileName);
        
        if (!File.Exists(filePath))
            Assert.Fail($"Example file not found: {filePath}");

        var exception = Record.Exception(() => WorkBookReader.LoadFromFile(filePath));
        
        if (exception != null)
            Assert.Fail($"Failed to load {fileName}: {exception.Message}\n{exception.StackTrace}");
        
        var workbook = WorkBookReader.LoadFromFile(filePath);
        Assert.NotNull(workbook);
        Assert.NotEmpty(workbook.Sheets);
    }

    [Fact]
    public void AllRecentExampleFiles_Exist()
    {
        var expectedFiles = new[]
        {
            "31_RowHeight.xlsx",
            "32_FreezePanes.xlsx",
            "33_TextAlignment.xlsx",
            "34_Hyperlinks.xlsx",
            "35_ReadExcel.xlsx",
            "36_RoundTripEditing.xlsx"
        };

        foreach (var fileName in expectedFiles)
        {
            var filePath = Path.Combine(ExamplesPath, fileName);
            Assert.True(File.Exists(filePath), $"Example file missing: {fileName}");
        }
    }
}
