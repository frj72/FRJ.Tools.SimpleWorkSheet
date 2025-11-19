using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class GenericTableBasicExample : IExample
{
    public string Name => "74_GenericTableBasic";
    public string Description => "Basic Generic Table import";

    public void Run()
    {
        var table = GenericTable.Create("Name", "Age", "Department", "Salary");
        
        table.AddRow(new CellValue("Alice"), new CellValue(28), new CellValue("Engineering"), new CellValue(75000m));
        table.AddRow(new CellValue("Bob"), new CellValue(35), new CellValue("Sales"), new CellValue(65000m));
        table.AddRow(new CellValue("Charlie"), new CellValue(42), new CellValue("Management"), new CellValue(95000m));
        table.AddRow(new CellValue("Diana"), new CellValue(31), new CellValue("Engineering"), new CellValue(80000m));
        
        var sheet = WorksheetBuilder.FromGenericTable(table)
            .WithSheetName("Employees")
            .WithHeaderStyle(style => style
                .WithFont(font => font.Bold().WithSize(12))
                .WithFillColor("FFE0E0E0"))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, $"{Name}.xlsx");
    }
}
