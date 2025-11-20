using System.Globalization;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using SkiaSharp;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class CellExtensions
{
    private const double DefaultWidth = 65.0 / 6.0;

    extension(Cell cell)
    {
        public Cell SetValue(CellValue value) => cell with { Value = value };

        public Cell SetFont(CellFont font)
        {
            if (!font.HasValidColor())
                throw new ArgumentException("Invalid font color format", nameof(font));
            return cell with { Style = cell.Style.WithFont(font) };
        }

        public Cell SetColor(string color)
        {
            if (!color.IsValidColor())
                throw new ArgumentException("Invalid color format", nameof(color));
            return cell with { Style = cell.Style.WithFillColor(color) };
        }

        public Cell SetBorders(CellBorders borders)
        {
            if (!borders.HasValidColors())
                throw new ArgumentException("Invalid border color format", nameof(borders));
            return cell with { Style = cell.Style.WithBorders(borders) };
        }
        public Cell SetDefaultFormatting()
            => cell with { Style = WorkSheetDefaults.DefaultCellStyle };
    }

    public static double EstimateMaxWidth(this IEnumerable<Cell> cells)
    {
        var maxWidth = -1.0;
        foreach (var cell in cells)
        {
            var cellFont = cell.Font ?? WorkSheetDefaults.Font;
            var fontStyleWeight = cellFont.Bold
                ? SKFontStyleWeight.Bold
                : SKFontStyleWeight.Normal;
            var fontStyleSlant = cellFont.Italic
                ? SKFontStyleSlant.Italic
                : SKFontStyleSlant.Upright;
            using var typeface = SKTypeface.FromFamilyName(
                cellFont.Name,
                fontStyleWeight,
                SKFontStyleWidth.Normal,
                fontStyleSlant
            );

            var textSize = (float)(cellFont.Size ?? WorkSheetDefaults.FontSize);

            using var font = new SKFont(typeface, textSize);


            var textToEstimate = cell.Value.CellValueType() switch
            {
                CellValueBasicType.String => cell.Value.AsString(),
                CellValueBasicType.FloatingPointNumber =>
                    cell.Value.AsDecimal().ToString("F3", CultureInfo.InvariantCulture) + " ",
                CellValueBasicType.IntegerNumber =>
                    cell.Value.AsLong().ToString(CultureInfo.InvariantCulture) + " ",
                CellValueBasicType.DateType =>
                    cell.Value.AsDateTime().ToString("yyyy-MM-dd HH:mm:ss"),
                _ => cell.Value.AsString()
            };

            var rawWidthPx = font.MeasureText(textToEstimate);
            var width = Math.Ceiling(rawWidthPx / 6.0);

            if (width > maxWidth)
                maxWidth = width;
        }

        return maxWidth <= 0D
            ? DefaultWidth
            : maxWidth;
    }
}