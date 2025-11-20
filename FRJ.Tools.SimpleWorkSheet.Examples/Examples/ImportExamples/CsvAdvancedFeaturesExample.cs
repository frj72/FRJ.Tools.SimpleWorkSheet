using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class CsvAdvancedFeaturesExample : IExample
{
    public string Name => "CSV with Advanced Features";
    public string Description => "CSV import with styling, column ordering, filtering, formatting, and conditional styles";

    public void Run()
    {
        var csvPath = Path.Combine("Resources", "Data", "Csv", "MixedTypes.csv");
        
        var sheet = WorksheetBuilder.FromCsvFile(csvPath)
            .WithSheetName("CSV Advanced")
            .WithColumnOrder("description", "amount", "isActive", "signupDate")
            .WithExcludeColumns("id")
            .WithHeaderStyle(style => style
                .WithFillColor("ED7D31")
                .WithFont(f => f.Bold().WithColor("FFFFFF")))
            .WithNumberFormat("amount", NumberFormat.Float2)
            .WithDateFormat(DateFormat.IsoDate)
            .WithConditionalStyle("isActive",
                val => val.IsString() && val.Value.AsT2 == "True",
                style => style.WithFillColor("C6E0B4"))
            .AutoFitAllColumns()
            .Build();
        
        ExampleRunner.SaveWorkSheet(sheet, "72_CsvAdvancedFeatures.xlsx");
    }
}
