using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.CalibrationExamples;

public class SimpleCalibrationWorkbook09Example : IExample
{
    public string Name => "Simple Calibration Workbook - 0.9 Factor";
    public string Description => "Demonstrates CreateSimpleCalibrationWorkBook with 0.9 calibration factor";

    public void Run()
    {
        var workbook = EnvironmentSheetInfo.CreateSimpleCalibrationWorkBook(0.9);
        ExampleRunner.SaveWorkBook(workbook, "110_SimpleCalibrationWorkbook_0.9.xlsx");
    }
}
