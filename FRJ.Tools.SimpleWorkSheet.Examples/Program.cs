using FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.BatchExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.CalibrationExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.IntegrationExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples;

public static class Program
{
    public static void Main()
    {
        var examples = new List<IExample>
        {
            new HelloWorldExample(1),
            new DataTypesExample(2),
            new SimpleTableExample(3),
            new CellPositioningExample(4),
            new ValueOnlyExample(5),
            new FontVariationsExample(6),
            new BackgroundColorsExample(7),
            new BorderStylesExample(8),
            new CompleteStylingExample(9),
            new RowAdditionExample(10),
            new ColumnAdditionExample(11),
            new BulkUpdatesExample(12),
            new MetadataTrackingExample(13),
            new ImportOptionsExample(14),
            new CsvSimulationExample(15),
            new ConditionalFormattingExample(16),
            new ReusableStylesExample(17),
            new ComplexTableExample(18),
            new InvoiceGeneratorExample(19),
            new DataVisualizationExample(20),
            new BasicFormulasExample(21),
            new SumFormulaExample(22),
            new AverageFormulaExample(23),
            new PercentageFormulaExample(24),
            new ConditionalFormulaExample(25),
            new MultiRangeFormulaExample(26),
            new CountFormulaExample(27),
            new MinMaxFormulaExample(28),
            new JsonArrayImportExample(29),
            new JsonSparseArrayImportExample(30),
            new RowHeightExample(31),
            new FreezePanesExample(32),
            new TextAlignmentExample(33),
            new HyperlinksExample(34),
            new CellMergingExample(37),
            new ReadExcelExample(35),
            new RoundTripEditingExample(36),
            new DataValidationExample(38),
            new NamedRangesExample(39),
            new AdvancedFormulasExample(40),
            new BarChartExample(41),
            new LineChartExample(42),
            new LineChartMultipleSeriesExample(89),
            new LineChartWithoutCategoriesExample(90),
            new LineChartMultipleSeriesWithoutCategoriesExample(91),
            new BarChartWithoutCategoriesExample(92),
            new AreaChartWithoutCategoriesExample(93),
            new PieChartExample(43),
            new ScatterChartExample(44),
            new ChartSheetExample(45),
            new AutoFitColumnsExample(46),
            new EnvironmentSheetInfoExample(108),
            new SheetTabColorsExample(47),
            new SheetVisibilityExample(48),
            new ExcelTablesExample(49),
            new InsertImagesExample(50),
            new AreaChartExample(51),
            new StackedAreaChartExample(52),
            new ChartFormattingExample(53),
            new ChartSeriesNamesExample(84),
            new JsonFluentImportPricesExample(54),
            new JsonFluentImportPricesFlatExample(55),
            new JsonFluentImportPersonsExample(56),
            new JsonArrayWithStylingExample(57),
            new JsonFlatObjectWithParsersExample(58),
            new JsonMultiColumnAllFeaturesExample(59),
            new JsonNestedObjectsExample(60),
            new JsonToWorkbookExample(61),
            new JsonToWorkbookLineChartExample(62),
            new JsonToWorkbookBarChartExample(63),
            new JsonToWorkbookMultipleChartsExample(64),
            new JsonColumnOrderingExample(65),
            new JsonColumnFilteringExample(66),
            new JsonDateNumberFormattingExample(67),
            new JsonConditionalStylingExample(68),
            new JsonAdvancedWorkbookExample(69),
            new CsvWithHeaderExample(70),
            new CsvWithoutHeaderExample(71),
            new CsvAdvancedFeaturesExample(72),
            new CsvToWorkbookWithChartExample(73),
            new GenericTableBasicExample(74),
            new GenericTableAdvancedExample(75),
            new GenericTableToWorkbookExample(76),
            new GenericTableWorkbookAdvancedExample(77),
            new BuiltInColorsExample(78),
            new ClassToWorkbookExample(79),
            new ChartSeriesNamesImportAutoExample(80),
            new ChartSeriesNamesImportCustomExample(81),
            new ChartSeriesNamesImportCsvExample(82),
            new ChartSeriesNamesImportGenericTableExample(83),
            new MultiTableMultiChartCommonSheetExample(85),
            new ColumnHidingExample(86),
            new RowHidingExample(87),
            new ConditionalHidingExample(88),
            new LineChartWithDataLabelsExample(94),
            new LineChartMultipleSeriesWithDataLabelsExample(95),
            new ScatterChartWithDataLabelsExample(96),
            new LineChartWithYAxisLabelsExample(97),
            new LineChartMultipleSeriesWithYAxisLabelsExample(98),
            new ScatterChartWithYAxisLabelsExample(99),
            new LineChartWithoutYAxisLabelsExample(100),
            new BarChartWithoutYAxisLabelsExample(101),
            new AreaChartWithoutYAxisLabelsExample(102),
            new LineChartWithCustomColorExample(103),
            new LineChartMultipleSeriesWithColorsExample(104),
            new BarChartWithCustomColorsExample(105),
            new AreaChartWithCustomColorsExample(106),
            new ChartColorPaletteExample(107),
            new SimpleCalibrationWorkbookDefaultExample(109),
            new SimpleCalibrationWorkbook09Example(110),
            new SimpleCalibrationWorkbook11Example(111)
        };

        Console.WriteLine("FRJ.Tools.SimpleWorkSheet - Examples");
        Console.WriteLine("====================================\n");

        RunAllExamples(examples);
    }

    private static void RunAllExamples(List<IExample> examples)
    {
        Console.WriteLine("Running all examples...\n");
        
        foreach (var example in examples) ExampleRunner.RunExample(example);

        Console.WriteLine($"\n✓ All {examples.Count} examples completed!");
        Console.WriteLine($"Output files are in: {Path.Combine(Directory.GetCurrentDirectory(), "Output")}");
    }
}
