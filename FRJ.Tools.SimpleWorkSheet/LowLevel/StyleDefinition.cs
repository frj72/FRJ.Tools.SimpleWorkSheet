using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.LowLevel;

public record StyleDefinition
{
    public string? FillColor { get; init; } 
    public CellFont Font { get; init; } = WorkSheetDefaults.Font;
    public CellBorders Borders { get; init; } = WorkSheetDefaults.CellBorders;
    public bool FormatAsDate { get; init; }
}

