using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.CalibrationExamples;

public class SimpleCalibrationWorkbookDefaultExample : IExample
{
    public string Name => "Simple Calibration Workbook - Default (Auto-Fit)";
    public string Description => "Demonstrates CreateSimpleCalibrationWorkBook with default calibration (null = auto-fit)";


    public int ExampleNumber { get; }

    public SimpleCalibrationWorkbookDefaultExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var workbook = EnvironmentSheetInfo.CreateSimpleCalibrationWorkBook();
        ExampleRunner.SaveWorkBook(workbook, $"{ExampleNumber:000}_SimpleCalibrationWorkbook_Default.xlsx");
    }
}
