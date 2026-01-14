using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class GenericTableToWorkbookExample : IExample
{
    public string Name => "076_GenericTableToWorkbook";
    public string Description => "Generic Table to Workbook with chart";

    public void Run()
    {
        var table = GenericTable.Create("Month", "Revenue", "Expenses", "Profit");
        
        table.AddRow(new CellValue("January"), new CellValue(45000m), new CellValue(32000m), new CellValue(13000m));
        table.AddRow(new CellValue("February"), new CellValue(52000m), new CellValue(35000m), new CellValue(17000m));
        table.AddRow(new CellValue("March"), new CellValue(48000m), new CellValue(33000m), new CellValue(15000m));
        table.AddRow(new CellValue("April"), new CellValue(61000m), new CellValue(38000m), new CellValue(23000m));
        table.AddRow(new CellValue("May"), new CellValue(58000m), new CellValue(36000m), new CellValue(22000m));
        table.AddRow(new CellValue("June"), new CellValue(67000m), new CellValue(40000m), new CellValue(27000m));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithWorkbookName("Financial Report")
            .WithDataSheetName("Monthly Data")
            .WithHeaderStyle(style => style
                .WithFont(font => font.Bold().WithSize(12))
                .WithFillColor("FF4472C4")
                .WithFont(font => font.WithColor(Colors.White)))
            .WithNumberFormat("Revenue", NumberFormat.Float2)
            .WithNumberFormat("Expenses", NumberFormat.Float2)
            .WithNumberFormat("Profit", NumberFormat.Float2)
            .WithConditionalStyle("Profit",
                value => value.Value.AsT0 > 20000,
                style => style.WithFillColor("FFC6EFCE"))
            .WithFreezeHeaderRow()
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkBook(workbook, $"{Name}.xlsx");
    }
}
