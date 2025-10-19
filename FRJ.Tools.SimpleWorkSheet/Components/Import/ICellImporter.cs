using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public interface ICellImporter
{
    Cell ImportCell(string rawValue, ImportOptions? options = null);
    
    IEnumerable<Cell> ImportRow(IEnumerable<string> values, ImportOptions? options = null);
    
    IEnumerable<IEnumerable<Cell>> ImportRows(IEnumerable<IEnumerable<string>> rows, ImportOptions? options = null);
}
