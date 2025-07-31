using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorkSheet.Components.Book;

public class WorkBook
{
    public WorkBook(string name, IEnumerable<WorkSheet> sheets)
    {
        Name = name;
        Sheets = sheets;
    }

    public string Name { get; }
    public IEnumerable<WorkSheet> Sheets { get; }

    public void SaveToFile(string fileName)
    {
        var bytes = SheetConverter.ToBinaryExcelFile(this);
        File.WriteAllBytes("fileName", bytes);
    }
}
