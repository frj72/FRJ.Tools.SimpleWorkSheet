using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public class PieChart : Chart
{
    public uint Explosion { get; private set; }
    public uint FirstSliceAngle { get; private set; }
    public CellRange? CategoriesRange { get; private set; }
    public CellRange? ValuesRange { get; private set; }

    private PieChart() : base(ChartType.Pie)
    {
    }

    public static PieChart Create() => new();

    public PieChart WithTitle(string? title)
    {
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Title cannot be null or empty", nameof(title));

        Title = title;
        return this;
    }

    public PieChart WithDataRange(CellRange categoriesRange, CellRange valuesRange)
    {
        ChartDataRange.ValidateDataRange(categoriesRange);
        ChartDataRange.ValidateDataRange(valuesRange);

        CategoriesRange = categoriesRange;
        ValuesRange = valuesRange;
        return this;
    }

    public PieChart WithPosition(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        Position = new(fromColumn, fromRow, toColumn, toRow);
        return this;
    }

    public PieChart WithSize(int width, int height)
    {
        Size = new(width, height);
        return this;
    }

    public PieChart WithExplosion(uint percentage)
    {
        if (percentage > 100)
            throw new ArgumentException("Explosion percentage must be between 0 and 100", nameof(percentage));

        Explosion = percentage;
        return this;
    }

    public PieChart WithFirstSliceAngle(uint degrees)
    {
        if (degrees > 360)
            throw new ArgumentException("First slice angle must be between 0 and 360 degrees", nameof(degrees));

        FirstSliceAngle = degrees;
        return this;
    }

    public PieChart WithDataSourceSheet(string sheetName)
    {
        DataSourceSheet = sheetName;
        return this;
    }

    public new PieChart AddSeries(string name, CellRange dataRange)
    {
        base.AddSeries(name, dataRange);
        return this;
    }
}
