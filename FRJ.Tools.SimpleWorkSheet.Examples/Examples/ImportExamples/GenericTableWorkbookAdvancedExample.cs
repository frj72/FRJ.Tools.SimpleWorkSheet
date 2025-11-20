using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class GenericTableWorkbookAdvancedExample : IExample
{
    public string Name => "77_GenericTableWorkbookAdvanced";
    public string Description => "GenericTable to Workbook with all advanced features: column ordering, filtering, formatting, and conditional styles";

    public void Run()
    {
        var table = GenericTable.Create("EmployeeID", "Name", "Department", "Salary", "HireDate", "Status", "Email");
        
        table.AddRow(
            new CellValue(101),
            new CellValue("Alice Johnson"),
            new CellValue("Engineering"),
            new CellValue(85000m),
            new CellValue(new DateTime(2020, 3, 15)),
            new CellValue("Active"),
            new CellValue("alice@company.com"));
            
        table.AddRow(
            new CellValue(102),
            new CellValue("Bob Smith"),
            new CellValue("Sales"),
            new CellValue(62000m),
            new CellValue(new DateTime(2021, 7, 1)),
            new CellValue("Active"),
            new CellValue("bob@company.com"));
            
        table.AddRow(
            new CellValue(103),
            new CellValue("Carol Davis"),
            new CellValue("Engineering"),
            new CellValue(95000m),
            new CellValue(new DateTime(2019, 1, 10)),
            new CellValue("Active"),
            new CellValue("carol@company.com"));
            
        table.AddRow(
            new CellValue(104),
            new CellValue("David Wilson"),
            new CellValue("Marketing"),
            new CellValue(58000m),
            new CellValue(new DateTime(2022, 2, 20)),
            new CellValue("Active"),
            new CellValue("david@company.com"));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithWorkbookName("Employee Analysis")
            .WithDataSheetName("Employee Data")
            .WithColumnOrder("Name", "Department", "Salary", "HireDate", "EmployeeID")
            .WithExcludeColumns("Status", "Email")
            .WithHeaderStyle(style => style
                .WithFont(font => font.Bold().WithSize(12).WithColor(Colors.White))
                .WithFillColor("FF4472C4"))
            .WithDateFormat(DateFormat.IsoDate)
            .WithNumberFormat("Salary", NumberFormat.Float2)
            .WithConditionalStyle("Salary",
                value => value.Value.AsT0 > 80000,
                style => style.WithFillColor("FFC6EFCE"))
            .WithFreezeHeaderRow()
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, $"{Name}.xlsx");
    }
}
