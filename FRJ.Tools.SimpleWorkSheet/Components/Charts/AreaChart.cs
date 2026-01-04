using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public class AreaChart : Chart
{
    public bool Stacked { get; private set; }
    public CellRange? CategoriesRange { get; private set; }
    public CellRange? ValuesRange { get; private set; }

    private AreaChart() : base(ChartType.Area)
    {
    }

    public static AreaChart Create() => new();

    public AreaChart WithTitle(string? title)
    {
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Title cannot be null or empty", nameof(title));

        Title = title;
        return this;
    }

    public AreaChart WithDataRange(CellRange categoriesRange, CellRange valuesRange)
    {
        ChartDataRange.ValidateDataRange(categoriesRange);
        ChartDataRange.ValidateDataRange(valuesRange);

        CategoriesRange = categoriesRange;
        ValuesRange = valuesRange;
        return this;
    }

    public AreaChart WithPosition(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        Position = new(fromColumn, fromRow, toColumn, toRow);
        return this;
    }

    public AreaChart WithSize(int width, int height)
    {
        Size = new(width, height);
        return this;
    }

    public AreaChart WithStacked(bool stacked)
    {
        Stacked = stacked;
        return this;
    }

    public AreaChart WithDataSourceSheet(string sheetName)
    {
        DataSourceSheet = sheetName;
        return this;
    }

    public new AreaChart AddSeries(string name, CellRange dataRange)
    {
        base.AddSeries(name, dataRange);
        return this;
    }

    public AreaChart WithLegendPosition(ChartLegendPosition position)
    {
        LegendPosition = position;
        return this;
    }

    public AreaChart WithCategoryAxisTitle(string title)
    {
        CategoryAxisTitle = title;
        return this;
    }

    public AreaChart WithValueAxisTitle(string title)
    {
        ValueAxisTitle = title;
        return this;
    }

    public AreaChart WithDataLabels(bool show)
    {
        ShowDataLabels = show;
        return this;
    }

    public AreaChart WithMajorGridlines(bool show)
    {
        ShowMajorGridlines = show;
        return this;
    }

    public AreaChart WithSeriesName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Series name cannot be null or empty", nameof(name));

        SingleSeriesName = name;
        return this;
    }
}
