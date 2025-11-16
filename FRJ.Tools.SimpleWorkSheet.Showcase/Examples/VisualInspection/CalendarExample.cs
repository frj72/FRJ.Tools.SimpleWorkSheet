using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.VisualInspection;

public class CalendarExample : IShowcase
{
    public string Name => "Monthly Calendar";
    public string Description => "Calendar layout with merged cells";
    public string Category => "Visual Inspection";

    public void Run()
    {
        var sheet = new WorkSheet("Calendar");
        
        sheet.AddCell(new(0, 0), "November 2025", cell => cell
            .WithFont(f => f.Bold().WithSize(16))
            .WithStyle(s => s.WithHorizontalAlignment(HorizontalAlignment.Center)));
        sheet.MergeCells(0, 0, 6, 0);
        
        var days = new[] { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
        for (uint col = 0; col < 7; col++)
            sheet.AddCell(new(col, 1), days[col], cell => cell
                .WithFont(f => f.Bold())
                .WithColor("4472C4")
                .WithStyle(s => s.WithHorizontalAlignment(HorizontalAlignment.Center)));
        
        var firstDay = new DateTime(2025, 11, 1);
        var startDayOfWeek = (int)firstDay.DayOfWeek;
        var daysInMonth = DateTime.DaysInMonth(2025, 11);
        
        uint row = 2;
        var day = 1;
        
        for (var week = 0; week < 6 && day <= daysInMonth; week++)
        {
            for (uint col = 0; col < 7; col++)
                if (week == 0 && col < startDayOfWeek)
                    sheet.AddCell(new(col, row), "", cell => cell.WithColor("F0F0F0"));
                else if (day <= daysInMonth)
                {
                    sheet.AddCell(new(col, row), day, cell => cell
                        .WithStyle(s => s
                            .WithHorizontalAlignment(HorizontalAlignment.Center)
                            .WithVerticalAlignment(VerticalAlignment.Top)));
                    day++;
                }

            row++;
        }
        
        for (uint col = 0; col < 7; col++)
            sheet.SetColumnWith(col, 12.0);
        
        for (uint r = 2; r < 8; r++)
            sheet.SetRowHeight(r, 60.0);
        
        var workbook = new WorkBook("Calendar", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_14_Calendar.xlsx");
    }
}
