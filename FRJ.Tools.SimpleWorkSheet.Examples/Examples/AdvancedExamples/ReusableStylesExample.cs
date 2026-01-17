using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class ReusableStylesExample : IExample
{
    public string Name => "Reusable Styles";
    public string Description => "Creating and applying style templates";


    public int ExampleNumber { get; }

    public ReusableStylesExample(int exampleNumber) => ExampleNumber = exampleNumber;
    public void Run()
    {
        var sheet = new WorkSheet("ReusableStyles");
        
        var headerStyle = CellStyle.Create(
            fillColor: "4472C4",
            font: CellFont.Create(14, "Calibri", "FFFFFF", bold: true),
            borders: null,
            formatCode: null);
        
        var dataStyle = CellStyle.Create(
            fillColor: "F0F0F0",
            font: CellFont.Create(11, "Calibri", "000000"),
            borders: null,
            formatCode: null);
        
        var highlightStyle = CellStyle.Create(
            fillColor: "FFFF00",
            font: CellFont.Create(11, "Calibri", "000000", bold: true),
            borders: null,
            formatCode: null);
        
        sheet.AddStyledCell(0, 0, "Name", headerStyle);
        sheet.AddStyledCell(1, 0, "Department", headerStyle);
        sheet.AddStyledCell(2, 0, "Salary", headerStyle);
        
        sheet.AddStyledCell(0, 1, "John", dataStyle);
        sheet.AddStyledCell(1, 1, "Engineering", dataStyle);
        sheet.AddStyledCell(2, 1, "$100,000", highlightStyle);
        
        sheet.AddStyledCell(0, 2, "Jane", dataStyle);
        sheet.AddStyledCell(1, 2, "Marketing", dataStyle);
        sheet.AddStyledCell(2, 2, "$90,000", dataStyle);
        
        ExampleRunner.SaveWorkSheet(sheet, $"{ExampleNumber:000}_ReusableStyles.xlsx");
    }
}