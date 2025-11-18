using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.AdvancedExamples;

public class InsertImagesExample : IExample
{
    public string Name => "Insert Images";
    public string Description => "Adding images to worksheets";

    public void Run()
    {
        var sheet = new WorkSheet("ImageDemo");

        sheet.AddCell(new(0, 0), "Company Report with Images", cell => cell
            .WithFont(font => font.Bold().WithSize(18))
            .WithStyle(style => style.WithFillColor("4472C4")));

        sheet.AddCell(new(0, 2), "PNG Format:", cell => cell.WithFont(font => font.Bold()));
        sheet.AddCell(new(0, 20), "JPEG Format:", cell => cell.WithFont(font => font.Bold()));

        const string pngPath = "Resources/Images/Wikipedia_Logo_1.0.png";
        const string jpegPath = "Resources/Images/Wikipedia_logo_593.jpg";

        if (File.Exists(pngPath))
            sheet.AddImage(pngPath, new(0, 3), 300, 300);

        if (File.Exists(jpegPath))
            sheet.AddImage(jpegPath, new(0, 21), 200, 200);

        sheet.AddCell(new(6, 2), "Financial Summary", cell => cell
            .WithFont(font => font.Bold().WithSize(14)));
        
        sheet.AddCell(new(6, 4), "Quarter");
        sheet.AddCell(new(7, 4), "Revenue");
        sheet.AddCell(new(8, 4), "Expenses");
        sheet.AddCell(new(9, 4), "Profit");
        
        sheet.AddCell(new(6, 5), "Q1");
        sheet.AddCell(new(7, 5), 1250000);
        sheet.AddCell(new(8, 5), 890000);
        sheet.AddCell(new(9, 5), 360000);
        
        sheet.AddCell(new(6, 6), "Q2");
        sheet.AddCell(new(7, 6), 1380000);
        sheet.AddCell(new(8, 6), 920000);
        sheet.AddCell(new(9, 6), 460000);
        
        sheet.AutoFitAllColumns();

        ExampleRunner.SaveWorkSheet(sheet, "50_InsertImages.xlsx");
    }
}
