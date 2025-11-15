using DocumentFormat.OpenXml.Spreadsheet;
using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using Cell = FRJ.Tools.SimpleWorkSheet.Components.SimpleCell.Cell;
// ReSharper disable UnusedMember.Global

namespace FRJ.Tools.SimpleWorkSheet.LowLevel;

public class NumberingFormatsProvider
{
    private readonly List<NumberingFormat> _internalNumberingFormats =
    [ 
        new() { NumberFormatId = DateFormat.IsoDateTime.ToExcelFormatId(), FormatCode = DateFormat.IsoDateTime.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.IsoDate.ToExcelFormatId(), FormatCode = DateFormat.IsoDate.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.DateTime.ToExcelFormatId(), FormatCode = DateFormat.DateTime.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.DateOnly.ToExcelFormatId(), FormatCode = DateFormat.DateOnly.ToExcelFormatString() },
        new() { NumberFormatId = DateFormat.TimeOnly.ToExcelFormatId(), FormatCode = DateFormat.TimeOnly.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Integer.ToExcelFormatId(), FormatCode = NumberFormat.Integer.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Float2.ToExcelFormatId(), FormatCode = NumberFormat.Float2.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Float3.ToExcelFormatId(), FormatCode = NumberFormat.Float3.ToExcelFormatString() },
        new() { NumberFormatId = NumberFormat.Float4.ToExcelFormatId(), FormatCode = NumberFormat.Float4.ToExcelFormatString() }
    ];
    
    public NumberingFormats NumberingFormats()
    {
        NumberingFormats numberingFormats = new() { Count = (uint)_internalNumberingFormats.Count };
        numberingFormats.Append(_internalNumberingFormats);
        return numberingFormats;
    }

    public uint GetOrCreateNumberFormatId(string formatCode)
    {
        var existing = _internalNumberingFormats.FirstOrDefault(nf => nf.FormatCode == formatCode);
        if (existing != null && existing.NumberFormatId != null)
            return existing.NumberFormatId;

        var newId = (uint)_internalNumberingFormats.Count + 164;
        _internalNumberingFormats.Add(new() { NumberFormatId = newId, FormatCode = formatCode });
        return newId;
    }

    public uint GetOrCreateNumberFormatId(Cell cell)
    {
        if (cell.FormatCode is not null)
            return GetOrCreateNumberFormatId(cell.FormatCode);
        
        if (cell.Value.IsDateTime() || cell.Value.IsDateTimeOffset())
            return DateFormat.IsoDateTime.ToExcelFormatId();

        if (cell.Value.IsDecimal())
            return NumberFormat.Float2.ToExcelFormatId();

        return cell.Value.IsLong() ? NumberFormat.Integer.ToExcelFormatId() : 0;
    }
}