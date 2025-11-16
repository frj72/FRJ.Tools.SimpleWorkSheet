using FRJ.Tools.SimpleWorkSheet.Showcase;
using FRJ.Tools.SimpleWorkSheet.Showcase.Examples.EdgeCases;
using FRJ.Tools.SimpleWorkSheet.Showcase.Examples.StressTests;
using FRJ.Tools.SimpleWorkSheet.Showcase.Examples.VisualInspection;
using FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ChartShowcase;
using FRJ.Tools.SimpleWorkSheet.Showcase.Examples.ValidationShowcase;

Console.WriteLine("====================================");
Console.WriteLine("SimpleWorkSheet Showcase");
Console.WriteLine("====================================");
Console.WriteLine();
Console.WriteLine("This will generate 24 Excel files for manual inspection.");
Console.WriteLine("Files will be saved to: ./Output/");
Console.WriteLine();

var tests = new List<IShowcase>
{
    new LargeDatasetExample(),
    new MaximumColumnsExample(),
    new UnicodeExample(),
    new BorderStylesMatrixExample(),
    new ColorPaletteExample(),
    new ExtremeSizesExample(),
    
    new HundredThousandCellsExample(),
    new FiftySheetsExample(),
    new ComplexFormulasExample(),
    new LargeChartExample(),
    
    new FinancialReportExample(),
    new CalendarExample(),
    new GradebookExample(),
    new InventoryExample(),
    new BudgetExample(),
    new MultiLevelHeaderExample(),
    
    new MultiChartDashboardExample(),
    new AllChartTypesExample(),
    new ChartFormattingExample(),
    
    new AllValidationTypesExample(),
    new FormInputExample(),
    new DateRangeValidationExample()
};

var runAll = args.Length == 0 || (args.Length > 0 && args[0] == "all");
var testNumber = -1;

if (!runAll && args.Length > 0)
{
    if (int.TryParse(args[0], out var num) && num >= 1 && num <= tests.Count)
        testNumber = num - 1;
    else
    {
        Console.WriteLine($"Invalid test number. Please specify 1-{tests.Count} or 'all'");
        return;
    }
}

var startTime = DateTime.Now;

if (runAll)
{
    Console.WriteLine($"Running all {tests.Count} showcase examples...\n");
    
    for (var i = 0; i < tests.Count; i++)
    {
        Console.WriteLine($"[{i + 1}/{tests.Count}] Running {tests[i].Category}: {tests[i].Name}");
        ShowcaseRunner.RunShowcase(tests[i]);
    }
}
else
{
    var test = tests[testNumber];
    Console.WriteLine($"Running test {testNumber + 1}: {test.Name}\n");
    ShowcaseRunner.RunShowcase(test);
}

var elapsed = DateTime.Now - startTime;
Console.WriteLine();
Console.WriteLine("====================================");
Console.WriteLine($"Completed in {elapsed.TotalSeconds:F2} seconds");
Console.WriteLine("====================================");
Console.WriteLine();
Console.WriteLine("Manual Review Instructions:");
Console.WriteLine("1. Open each generated .xlsx file in Excel");
Console.WriteLine("2. Verify all data displays correctly");
Console.WriteLine("3. Test any interactive features (validations, hyperlinks)");
Console.WriteLine("4. Check chart rendering");
Console.WriteLine("5. Verify formatting appears as expected");
Console.WriteLine();
Console.WriteLine("Files location: ./Output/");
