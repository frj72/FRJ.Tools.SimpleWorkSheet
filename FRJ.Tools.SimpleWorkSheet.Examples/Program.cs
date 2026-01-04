using FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.BasicExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.BatchExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.FormulaExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.IntegrationExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples;

public static class Program
{
    public static void Main(string[] args)
    {
        var examples = new List<IExample>
        {
            new HelloWorldExample(),
            new DataTypesExample(),
            new SimpleTableExample(),
            new CellPositioningExample(),
            new ValueOnlyExample(),
            new FontVariationsExample(),
            new BackgroundColorsExample(),
            new BorderStylesExample(),
            new CompleteStylingExample(),
            new RowAdditionExample(),
            new ColumnAdditionExample(),
            new BulkUpdatesExample(),
            new MetadataTrackingExample(),
            new ImportOptionsExample(),
            new CsvSimulationExample(),
            new ConditionalFormattingExample(),
            new ReusableStylesExample(),
            new ComplexTableExample(),
            new InvoiceGeneratorExample(),
            new DataVisualizationExample(),
            new BasicFormulasExample(),
            new SumFormulaExample(),
            new AverageFormulaExample(),
            new PercentageFormulaExample(),
            new ConditionalFormulaExample(),
            new MultiRangeFormulaExample(),
            new CountFormulaExample(),
            new MinMaxFormulaExample(),
            new JsonArrayImportExample(),
            new JsonSparseArrayImportExample(),
            new RowHeightExample(),
            new FreezePanesExample(),
            new TextAlignmentExample(),
            new HyperlinksExample(),
            new CellMergingExample(),
            new ReadExcelExample(),
            new RoundTripEditingExample(),
            new DataValidationExample(),
            new NamedRangesExample(),
            new AdvancedFormulasExample(),
            new BarChartExample(),
            new LineChartExample(),
            new PieChartExample(),
            new ScatterChartExample(),
            new ChartSheetExample(),
            new AutoFitColumnsExample(),
            new SheetTabColorsExample(),
            new SheetVisibilityExample(),
            new ExcelTablesExample(),
            new InsertImagesExample(),
            new AreaChartExample(),
            new StackedAreaChartExample(),
            new ChartFormattingExample(),
            new ChartSeriesNamesExample(),
            new JsonFluentImportPricesExample(),
            new JsonFluentImportPricesFlatExample(),
            new JsonFluentImportPersonsExample(),
            new JsonArrayWithStylingExample(),
            new JsonFlatObjectWithParsersExample(),
            new JsonMultiColumnAllFeaturesExample(),
            new JsonNestedObjectsExample(),
            new JsonToWorkbookExample(),
            new JsonToWorkbookLineChartExample(),
            new JsonToWorkbookBarChartExample(),
            new JsonToWorkbookMultipleChartsExample(),
            new JsonColumnOrderingExample(),
            new JsonColumnFilteringExample(),
            new JsonDateNumberFormattingExample(),
            new JsonConditionalStylingExample(),
            new JsonAdvancedWorkbookExample(),
            new CsvWithHeaderExample(),
            new CsvWithoutHeaderExample(),
            new CsvAdvancedFeaturesExample(),
            new CsvToWorkbookWithChartExample(),
            new GenericTableBasicExample(),
            new GenericTableAdvancedExample(),
            new GenericTableToWorkbookExample(),
            new GenericTableWorkbookAdvancedExample(),
            new BuiltInColorsExample(),
            new ClassToWorkbookExample(),
            new ChartSeriesNamesImportExample(),
            new MultiTableMultiChartCommonSheetExample()
        };

        Console.WriteLine("FRJ.Tools.SimpleWorkSheet - Examples");
        Console.WriteLine("====================================\n");

        if (args.Length > 0 && args[0] == "all")
        {
            RunAllExamples(examples);
            return;
        }

        ShowMenu(examples);
    }

    private static void ShowMenu(List<IExample> examples)
    {
        while (true)
        {
            Console.WriteLine("\nSelect an example to run:");
            Console.WriteLine("0. Run all examples");
            
            for (var i = 0; i < examples.Count; i++) Console.WriteLine($"{i + 1}. {examples[i].Name}");

            Console.WriteLine("q. Quit");
            Console.Write("\nYour choice: ");
            
            var input = Console.ReadLine();
            
            if (input?.ToLower() == "q")
                break;
                
            if (int.TryParse(input, out var choice))
                switch (choice)
                {
                    case 0:
                        RunAllExamples(examples);
                        break;
                    case > 0 when choice <= examples.Count:
                        ExampleRunner.RunExample(examples[choice - 1]);
                        break;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            else
                Console.WriteLine("Invalid input. Please enter a number or 'q' to quit.");
        }
        
        Console.WriteLine("\nGoodbye!");
    }

    private static void RunAllExamples(List<IExample> examples)
    {
        Console.WriteLine("\nRunning all examples...\n");
        
        foreach (var example in examples) ExampleRunner.RunExample(example);

        Console.WriteLine($"\n✓ All {examples.Count} examples completed!");
        Console.WriteLine($"Output files are in: {Path.Combine(Directory.GetCurrentDirectory(), "Output")}");
    }
}
