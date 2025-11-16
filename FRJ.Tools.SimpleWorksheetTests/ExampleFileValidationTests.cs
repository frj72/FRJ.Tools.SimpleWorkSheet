using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ExampleFileValidationTests
{
    private static readonly string ExamplesPath = Path.Combine(
        Directory.GetCurrentDirectory(), 
        "../../../..",
        "FRJ.Tools.SimpleWorkSheet.Examples/Output");

    [Theory]
    [InlineData("01_HelloWorld.xlsx")]
    [InlineData("02_DataTypes.xlsx")]
    [InlineData("03_SimpleTable.xlsx")]
    [InlineData("04_CellPositioning.xlsx")]
    [InlineData("05_ValueOnly.xlsx")]
    [InlineData("06_FontVariations.xlsx")]
    [InlineData("07_BackgroundColors.xlsx")]
    [InlineData("08_BorderStyles.xlsx")]
    [InlineData("09_CompleteStyling.xlsx")]
    [InlineData("10_RowAddition.xlsx")]
    [InlineData("11_ColumnAddition.xlsx")]
    [InlineData("12_BulkUpdates.xlsx")]
    [InlineData("13_MetadataTracking.xlsx")]
    [InlineData("14_ImportOptions.xlsx")]
    [InlineData("15_CsvSimulation.xlsx")]
    [InlineData("16_ConditionalFormatting.xlsx")]
    [InlineData("17_ReusableStyles.xlsx")]
    [InlineData("18_ComplexTable.xlsx")]
    [InlineData("19_InvoiceGenerator.xlsx")]
    [InlineData("20_DataVisualization.xlsx")]
    [InlineData("21_BasicFormulas.xlsx")]
    [InlineData("22_SumFormula.xlsx")]
    [InlineData("23_AverageFormula.xlsx")]
    [InlineData("24_PercentageFormula.xlsx")]
    [InlineData("25_ConditionalFormula.xlsx")]
    [InlineData("26_MultiRangeFormula.xlsx")]
    [InlineData("27_CountFormula.xlsx")]
    [InlineData("28_MinMaxFormula.xlsx")]
    [InlineData("29_JsonArrayImport.xlsx")]
    [InlineData("30_JsonSparseArrayImport.xlsx")]
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
    public void AllExampleFiles_Exist()
    {
        const int expectedCount = 36;
        var actualFiles = Directory.GetFiles(ExamplesPath, "*.xlsx");
        
        Assert.Equal(expectedCount, actualFiles.Length);
    }
}
