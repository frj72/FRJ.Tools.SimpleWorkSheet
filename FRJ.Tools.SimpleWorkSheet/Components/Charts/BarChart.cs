using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public class BarChart : Chart
{
    public BarChartOrientation Orientation { get; private set; } = BarChartOrientation.Vertical;
    public CellRange? CategoriesRange { get; private set; }
    public CellRange? ValuesRange { get; private set; }

    private BarChart() : base(ChartType.Bar)
    {
    }

    public static BarChart Create() => new();

    public BarChart WithTitle(string? title)
    {
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Title cannot be null or empty", nameof(title));

        Title = title;
        return this;
    }

    public BarChart WithDataRange(CellRange categoriesRange, CellRange valuesRange)
    {
        ChartDataRange.ValidateDataRange(categoriesRange);
        ChartDataRange.ValidateDataRange(valuesRange);

        CategoriesRange = categoriesRange;
        ValuesRange = valuesRange;
        return this;
    }

    public BarChart WithDataRange(CellRange valuesRange)
    {
        ChartDataRange.ValidateDataRange(valuesRange);

        CategoriesRange = null;
        ValuesRange = valuesRange;
        return this;
    }

    public BarChart WithPosition(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        Position = new(fromColumn, fromRow, toColumn, toRow);
        return this;
    }

    public BarChart WithSize(int width, int height)
    {
        Size = new(width, height);
        return this;
    }

    public BarChart WithOrientation(BarChartOrientation orientation)
    {
        Orientation = orientation;
        return this;
    }

    public BarChart WithDataSourceSheet(string sheetName)
    {
        DataSourceSheet = sheetName;
        return this;
    }

    public new BarChart AddSeries(string name, CellRange dataRange)
    {
        base.AddSeries(name, dataRange);
        return this;
    }

    public BarChart WithLegendPosition(ChartLegendPosition position)
    {
        LegendPosition = position;
        return this;
    }

    public BarChart WithCategoryAxisTitle(string title)
    {
        CategoryAxisTitle = title;
        return this;
    }

    public BarChart WithValueAxisTitle(string title)
    {
        ValueAxisTitle = title;
        return this;
    }

    public BarChart WithDataLabels(bool show)
    {
        ShowDataLabels = show;
        return this;
    }

    public BarChart WithMajorGridlines(bool show)
    {
        ShowMajorGridlines = show;
        return this;
    }

    public BarChart WithSeriesName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Series name cannot be null or empty", nameof(name));

        SingleSeriesName = name;
        return this;
    }
}
