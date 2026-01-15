using System.Text.RegularExpressions;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public partial class ExampleFileValidationTests
{
    private static readonly string ExamplesPath = Path.Combine(
        Directory.GetCurrentDirectory(),
        "../../../..",
        "FRJ.Tools.SimpleWorkSheet.Examples/Output");

    [Theory]
    [InlineData("001_HelloWorld.xlsx")]
    [InlineData("002_DataTypes.xlsx")]
    [InlineData("003_SimpleTable.xlsx")]
    [InlineData("004_CellPositioning.xlsx")]
    [InlineData("005_ValueOnly.xlsx")]
    [InlineData("006_FontVariations.xlsx")]
    [InlineData("007_BackgroundColors.xlsx")]
    [InlineData("008_BorderStyles.xlsx")]
    [InlineData("009_CompleteStyling.xlsx")]
    [InlineData("010_RowAddition.xlsx")]
    [InlineData("011_ColumnAddition.xlsx")]
    [InlineData("012_BulkUpdates.xlsx")]
    [InlineData("013_MetadataTracking.xlsx")]
    [InlineData("014_ImportOptions.xlsx")]
    [InlineData("015_CsvSimulation.xlsx")]
    [InlineData("016_ConditionalFormatting.xlsx")]
    [InlineData("017_ReusableStyles.xlsx")]
    [InlineData("018_ComplexTable.xlsx")]
    [InlineData("019_InvoiceGenerator.xlsx")]
    [InlineData("020_DataVisualization.xlsx")]
    [InlineData("021_BasicFormulas.xlsx")]
    [InlineData("022_SumFormula.xlsx")]
    [InlineData("023_AverageFormula.xlsx")]
    [InlineData("024_PercentageFormula.xlsx")]
    [InlineData("025_ConditionalFormula.xlsx")]
    [InlineData("026_MultiRangeFormula.xlsx")]
    [InlineData("027_CountFormula.xlsx")]
    [InlineData("028_MinMaxFormula.xlsx")]
    [InlineData("029_JsonArrayImport.xlsx")]
    [InlineData("030_JsonSparseArrayImport.xlsx")]
    [InlineData("031_RowHeight.xlsx")]
    [InlineData("032_FreezePanes.xlsx")]
    [InlineData("033_TextAlignment.xlsx")]
    [InlineData("034_Hyperlinks.xlsx")]
    [InlineData("035_ReadExcel.xlsx")]
    [InlineData("036_RoundTripEditing.xlsx")]
    [InlineData("037_CellMerging.xlsx")]
    [InlineData("038_DataValidation.xlsx")]
    [InlineData("039_NamedRanges.xlsx")]
    [InlineData("040_AdvancedFormulas.xlsx")]
    [InlineData("041_BarChart.xlsx")]
    [InlineData("042_LineChart.xlsx")]
    [InlineData("043_PieChart.xlsx")]
    [InlineData("044_ScatterChart.xlsx")]
    [InlineData("045_ChartSheet.xlsx")]
    [InlineData("046_AutoFitColumns.xlsx")]
    [InlineData("047_SheetTabColors.xlsx")]
    [InlineData("048_SheetVisibility.xlsx")]
    [InlineData("049_ExcelTables.xlsx")]
    [InlineData("050_InsertImages.xlsx")]
    [InlineData("051_AreaChart.xlsx")]
    [InlineData("052_StackedAreaChart.xlsx")]
    [InlineData("053_ChartFormatting.xlsx")]
    [InlineData("054_JsonArrayFluentImport.xlsx")]
    [InlineData("055_JsonFlatObjectImport.xlsx")]
    [InlineData("056_JsonMultiColumnArrayImport.xlsx")]
    [InlineData("057_JsonArrayWithStyling.xlsx")]
    [InlineData("058_JsonFlatObjectWithParsers.xlsx")]
    [InlineData("059_JsonMultiColumnAllFeatures.xlsx")]
    [InlineData("060_JsonNestedObjects.xlsx")]
    [InlineData("061_JsonToWorkbook.xlsx")]
    [InlineData("062_JsonToWorkbookLineChart.xlsx")]
    [InlineData("063_JsonToWorkbookBarChart.xlsx")]
    [InlineData("064_JsonToWorkbookMultipleCharts.xlsx")]
    [InlineData("065_JsonColumnOrdering.xlsx")]
    [InlineData("066_JsonColumnFiltering.xlsx")]
    [InlineData("067_JsonDateNumberFormatting.xlsx")]
    [InlineData("068_JsonConditionalStyling.xlsx")]
    [InlineData("069_JsonAdvancedWorkbook.xlsx")]
    [InlineData("070_CsvWithHeader.xlsx")]
    [InlineData("071_CsvWithoutHeader.xlsx")]
    [InlineData("072_CsvAdvancedFeatures.xlsx")]
    [InlineData("073_CsvToWorkbookWithChart.xlsx")]
    [InlineData("074_GenericTableBasic.xlsx")]
    [InlineData("075_GenericTableAdvanced.xlsx")]
    [InlineData("076_GenericTableToWorkbook.xlsx")]
    [InlineData("077_GenericTableWorkbookAdvanced.xlsx")]
    [InlineData("078_BuiltInColors.xlsx")]
    [InlineData("079_ClassToWorkbook.xlsx")]
    [InlineData("084_ChartSeriesNames.xlsx")]
    [InlineData("080_ChartSeriesNamesImport_Auto.xlsx")]
    [InlineData("081_ChartSeriesNamesImport_Custom.xlsx")]
    [InlineData("082_ChartSeriesNamesImport_CSV.xlsx")]
    [InlineData("083_ChartSeriesNamesImport_GenericTable.xlsx")]
    [InlineData("085_MultiTableMultiChartCommonSheet.xlsx")]
    [InlineData("086_ColumnHiding.xlsx")]
    [InlineData("087_RowHiding.xlsx")]
    [InlineData("088_ConditionalHiding.xlsx")]
    [InlineData("089_LineChartMultipleSeries.xlsx")]
    [InlineData("090_LineChartWithoutCategories.xlsx")]
    [InlineData("091_LineChartMultipleSeriesWithoutCategories.xlsx")]
    [InlineData("092_BarChartWithoutCategories.xlsx")]
    [InlineData("093_AreaChartWithoutCategories.xlsx")]
    [InlineData("094_LineChartWithDataLabels.xlsx")]
    [InlineData("095_LineChartMultipleSeriesWithDataLabels.xlsx")]
    [InlineData("096_ScatterChartWithDataLabels.xlsx")]
    [InlineData("097_LineChartWithExplicitYAxisLabels.xlsx")]
    [InlineData("098_LineChartMultipleSeriesWithExplicitYAxisLabels.xlsx")]
    [InlineData("099_ScatterChartWithExplicitYAxisLabels.xlsx")]
    [InlineData("100_LineChartWithoutYAxisLabels.xlsx")]
    [InlineData("101_BarChartWithoutYAxisLabels.xlsx")]
    [InlineData("102_AreaChartWithoutYAxisLabels.xlsx")]
    [InlineData("103_LineChartWithCustomColor.xlsx")]
    [InlineData("104_LineChartMultipleSeriesWithColors.xlsx")]
    [InlineData("105_BarChartWithCustomColors.xlsx")]
    [InlineData("106_AreaChartWithCustomColors.xlsx")]
    [InlineData("107_ChartColorPalette.xlsx")]
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
        const int expectedCount = 107;
        var actualFiles = Directory.GetFiles(ExamplesPath, "*.xlsx");

        Assert.Equal(expectedCount, actualFiles.Length);
    }

    [Fact]
    public void AllExampleFiles_Have_Correct_Numbering()
    {
        var prefixes = Directory.GetFiles(ExamplesPath, "*.xlsx")
            .Select(Path.GetFileName)
            .Where(name => !string.IsNullOrEmpty(name))
            .Cast<string>()
            .Select(fileName => ExampleFileName().Match(fileName))
            .Where(match => match.Success)
            .Select(match => int.Parse(match.Groups[1].Value))
            .ToList();
        
        var isValidSequence = prefixes.Order().SequenceEqual(Enumerable.Range(1, prefixes.Count));
        Assert.True(isValidSequence);
    }

    [GeneratedRegex(@"^(\d+)_[^\.]+\.xlsx$")]
    private static partial Regex ExampleFileName();
}
