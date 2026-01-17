using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.CalibrationExamples;

public class SimpleCalibrationWorkbook11Example : IExample
{
    public string Name => "Simple Calibration Workbook - 1.1 Factor";
    public string Description => "Demonstrates CreateSimpleCalibrationWorkBook with 1.1 calibration factor";

    public void Run()
    {
        var workbook = EnvironmentSheetInfo.CreateSimpleCalibrationWorkBook(1.1);
        ExampleRunner.SaveWorkBook(workbook, "111_SimpleCalibrationWorkbook_1.1.xlsx");
    }
}
