using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ImportWorkbookBuilderTests
{
    [Fact]
    public void FromJson_ValidJsonArray_ReturnsBuilder()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var builder = WorkbookBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void FromJson_ValidObject_ReturnsBuilder()
    {
        const string json = """{"a": 1, "b": 2}""";
        
        var builder = WorkbookBuilder.FromJson(json);
        
        Assert.NotNull(builder);
    }

    [Fact]
    public void Build_CreatesWorkbookWithDefaultName()
    {
        const string json = """[{"name": "John"}]""";
        
        var workbook = WorkbookBuilder.FromJson(json).Build();
        
        Assert.NotNull(workbook);
        Assert.Equal("Workbook", workbook.Name);
    }

    [Fact]
    public void Build_CreatesWorkbookWithSingleSheet()
    {
        const string json = """[{"name": "John"}]""";
        
        var workbook = WorkbookBuilder.FromJson(json).Build();
        
        Assert.Single(workbook.Sheets);
        Assert.Equal("Data", workbook.Sheets.First().Name);
    }

    [Fact]
    public void WithWorkbookName_SetsCustomName()
    {
        const string json = """[{"test": "value"}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithWorkbookName("CustomWorkbook")
            .Build();
        
        Assert.Equal("CustomWorkbook", workbook.Name);
    }

    [Fact]
    public void WithWorkbookName_EmptyString_ThrowsException()
    {
        const string json = """[{"test": "value"}]""";
        var builder = WorkbookBuilder.FromJson(json);
        
        Assert.Throws<ArgumentException>(() => builder.WithWorkbookName(""));
    }

    [Fact]
    public void WithDataSheetName_SetsCustomSheetName()
    {
        const string json = """[{"test": "value"}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithDataSheetName("CustomSheet")
            .Build();
        
        Assert.Equal("CustomSheet", workbook.Sheets.First().Name);
    }

    [Fact]
    public void WithDataSheetName_EmptyString_ThrowsException()
    {
        const string json = """[{"test": "value"}]""";
        var builder = WorkbookBuilder.FromJson(json);
        
        Assert.Throws<ArgumentException>(() => builder.WithDataSheetName(""));
    }

    [Fact]
    public void GetColumnIndexByName_ExistingColumn_ReturnsIndex()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var builder = WorkbookBuilder.FromJson(json);
        var index = builder.GetColumnIndexByName("age");
        
        Assert.NotNull(index);
        Assert.True(index >= 0);
    }

    [Fact]
    public void GetColumnIndexByName_NonExistingColumn_ReturnsNull()
    {
        const string json = """[{"name": "John"}]""";
        
        var builder = WorkbookBuilder.FromJson(json);
        var index = builder.GetColumnIndexByName("nonexistent");
        
        Assert.Null(index);
    }

    [Fact]
    public void GetColumnRangeByName_ExistingColumn_ReturnsRange()
    {
        const string json = """[{"name": "John"}, {"name": "Jane"}]""";
        
        var builder = WorkbookBuilder.FromJson(json);
        var range = builder.GetColumnRangeByName("name");
        
        Assert.NotNull(range);
        Assert.Equal(1u, range.Value.From.Y);
        Assert.Equal(2u, range.Value.To.Y);
    }

    [Fact]
    public void GetColumnRangeByName_NonExistingColumn_ReturnsNull()
    {
        const string json = """[{"name": "John"}]""";
        
        var builder = WorkbookBuilder.FromJson(json);
        var range = builder.GetColumnRangeByName("nonexistent");
        
        Assert.Null(range);
    }

    [Fact]
    public void Build_ChainedMethods_WorksCorrectly()
    {
        const string json = """[{"name": "John", "age": 30}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithWorkbookName("People")
            .WithDataSheetName("PersonData")
            .WithTrimWhitespace(true)
            .AutoFitAllColumns()
            .Build();
        
        Assert.Equal("People", workbook.Name);
        Assert.Equal("PersonData", workbook.Sheets.First().Name);
        Assert.NotEmpty(workbook.Sheets.First().ExplicitColumnWidths);
    }

    [Fact]
    public void GetColumnIndexByName_NestedProperty_FindsColumn()
    {
        const string json = """[{"user": {"name": "John"}}]""";
        
        var builder = WorkbookBuilder.FromJson(json);
        var index = builder.GetColumnIndexByName("user.name");
        
        Assert.NotNull(index);
        Assert.Equal(0, index.Value);
    }

    [Fact]
    public void GetColumnRangeByName_NestedProperty_ReturnsCorrectRange()
    {
        const string json = """[{"user": {"age": 30}}, {"user": {"age": 25}}]""";
        
        var builder = WorkbookBuilder.FromJson(json);
        var range = builder.GetColumnRangeByName("user.age");
        
        Assert.NotNull(range);
        Assert.Equal(0u, range.Value.From.X);
        Assert.Equal(0u, range.Value.To.X);
        Assert.Equal(1u, range.Value.From.Y);
        Assert.Equal(2u, range.Value.To.Y);
    }

    [Fact]
    public void WithPreserveOriginalValue_True_PreservesInMetadata()
    {
        const string json = """[{"price": 19.99}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithPreserveOriginalValue(true)
            .Build();
        
        var sheet = workbook.Sheets.First();
        var cell = sheet.Cells.Cells[new(0, 1)];
        Assert.NotNull(cell.Metadata);
        Assert.NotNull(cell.Metadata.OriginalValue);
        Assert.Equal("19.99", cell.Metadata.OriginalValue);
    }

    [Fact]
    public void WithPreserveOriginalValue_False_DoesNotPreserve()
    {
        const string json = """[{"price": 19.99}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithPreserveOriginalValue(false)
            .Build();
        
        var sheet = workbook.Sheets.First();
        var cell = sheet.Cells.Cells[new(0, 1)];
        Assert.Null(cell.Metadata);
    }

    [Fact]
    public void WithColumnParser_TransformsValues()
    {
        const string json = """[{"price": 100}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithColumnParser("price", value => new(value.Value.AsT0 * 2))
            .Build();
        
        var sheet = workbook.Sheets.First();
        var value = sheet.GetValue(0, 1);
        Assert.NotNull(value);
        Assert.Equal(200m, value.Value.AsT0);
    }

    [Fact]
    public void WithColumnParser_MultipleColumns_AllTransformed()
    {
        const string json = """[{"price": 100, "tax": 10}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithColumnParser("price", value => new(value.Value.AsT0 * 1.5m))
            .WithColumnParser("tax", value => new(value.Value.AsT0 * 2))
            .Build();
        
        var sheet = workbook.Sheets.First();
        Assert.Equal(150m, sheet.GetValue(0, 1)?.Value.AsT0);
        Assert.Equal(20m, sheet.GetValue(1, 1)?.Value.AsT0);
    }

    [Fact]
    public void WithChart_CreatesChartSheet()
    {
        const string json = """[{"month": "Jan", "sales": 1000}, {"month": "Feb", "sales": 1500}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("month", "sales")
                .AsLineChart())
            .Build();
        
        Assert.Equal(2, workbook.Sheets.Count());
        Assert.Equal("Data", workbook.Sheets.First().Name);
        Assert.Equal("Chart", workbook.Sheets.Last().Name);
    }

    [Fact]
    public void WithChart_OnSheet_SetsCustomChartSheetName()
    {
        const string json = """[{"x": 1, "y": 2}, {"x": 2, "y": 3}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .OnSheet("SalesChart")
                .UseColumns("x", "y")
                .AsLineChart())
            .Build();
        
        Assert.Equal("SalesChart", workbook.Sheets.Last().Name);
    }

    [Fact]
    public void WithChart_UseColumns_ResolvesColumnNames()
    {
        const string json = """[{"date": "2025-01-01", "price": 100}, {"date": "2025-01-02", "price": 110}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("date", "price")
                .AsLineChart())
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        Assert.NotEmpty(chartSheet.Charts);
    }

    [Fact]
    public void WithChart_AsLineChart_CreatesLineChart()
    {
        const string json = """[{"a": 1, "b": 2}, {"a": 2, "b": 3}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("a", "b")
                .AsLineChart())
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        Assert.Single(chartSheet.Charts);
        Assert.IsType<LineChart>(chartSheet.Charts[0]);
    }

    [Fact]
    public void WithChart_AsBarChart_CreatesBarChart()
    {
        const string json = """[{"category": "A", "value": 10}, {"category": "B", "value": 20}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("category", "value")
                .AsBarChart())
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        Assert.Single(chartSheet.Charts);
        Assert.IsType<BarChart>(chartSheet.Charts[0]);
    }

    [Fact]
    public void WithChart_AsAreaChart_CreatesAreaChart()
    {
        const string json = """[{"time": "Q1", "revenue": 5000}, {"time": "Q2", "revenue": 6000}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("time", "revenue")
                .AsAreaChart())
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        Assert.Single(chartSheet.Charts);
        Assert.IsType<AreaChart>(chartSheet.Charts[0]);
    }

    [Fact]
    public void WithChart_WithTitle_SetsChartTitle()
    {
        const string json = """[{"x": 1, "y": 2}, {"x": 2, "y": 4}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("x", "y")
                .AsLineChart()
                .WithTitle("Test Chart"))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var lineChart = (LineChart)chartSheet.Charts[0];
        Assert.Equal("Test Chart", lineChart.Title);
    }

    [Fact]
    public void WithChart_InvalidColumnName_ThrowsException()
    {
        const string json = """[{"a": 1}]""";
        
        var builder = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("nonexistent", "a")
                .AsLineChart());
        
        Assert.Throws<InvalidOperationException>(() => builder.Build());
    }

    [Fact]
    public void WithChart_MultipleCharts_CreatesMultipleSheets()
    {
        const string json = """[{"x": 1, "y": 2, "z": 3}, {"x": 2, "y": 4, "z": 6}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("x", "y")
                .AsLineChart())
            .WithChart(chart => chart
                .UseColumns("x", "z")
                .AsBarChart())
            .Build();
        
        Assert.Equal(3, workbook.Sheets.Count());
        Assert.Equal("Data", workbook.Sheets.ElementAt(0).Name);
        Assert.Equal("Chart", workbook.Sheets.ElementAt(1).Name);
        Assert.Equal("Chart", workbook.Sheets.ElementAt(2).Name);
    }

    [Fact]
    public void WithChart_AsPieChart_CreatesPieChart()
    {
        const string json = """[{"category": "A", "value": 30}, {"category": "B", "value": 70}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("category", "value")
                .AsPieChart())
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        Assert.Single(chartSheet.Charts);
        Assert.IsType<PieChart>(chartSheet.Charts[0]);
    }

    [Fact]
    public void WithChart_AsScatterChart_CreatesScatterChart()
    {
        const string json = """[{"x": 1.5, "y": 2.3}, {"x": 2.1, "y": 3.7}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("x", "y")
                .AsScatterChart())
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        Assert.Single(chartSheet.Charts);
        Assert.IsType<ScatterChart>(chartSheet.Charts[0]);
    }

    [Fact]
    public void WithChart_WithCategoryAxisTitle_SetsTitle()
    {
        const string json = """[{"month": "Jan", "sales": 100}, {"month": "Feb", "sales": 150}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("month", "sales")
                .AsLineChart()
                .WithCategoryAxisTitle("Month"))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var lineChart = (LineChart)chartSheet.Charts[0];
        Assert.Equal("Month", lineChart.CategoryAxisTitle);
    }

    [Fact]
    public void WithChart_WithValueAxisTitle_SetsTitle()
    {
        const string json = """[{"month": "Jan", "revenue": 1000}, {"month": "Feb", "revenue": 1200}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("month", "revenue")
                .AsAreaChart()
                .WithValueAxisTitle("Revenue ($)"))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var areaChart = (AreaChart)chartSheet.Charts[0];
        Assert.Equal("Revenue ($)", areaChart.ValueAxisTitle);
    }

    [Fact]
    public void WithChart_WithLegendPosition_SetsPosition()
    {
        const string json = """[{"cat": "A", "val": 10}, {"cat": "B", "val": 20}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("cat", "val")
                .AsBarChart()
                .WithLegendPosition(ChartLegendPosition.Bottom))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var barChart = (BarChart)chartSheet.Charts[0];
        Assert.Equal(ChartLegendPosition.Bottom, barChart.LegendPosition);
    }

    [Fact]
    public void WithChart_AllFormatting_AppliesCorrectly()
    {
        const string json = """[{"x": 1, "y": 100}, {"x": 2, "y": 200}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .OnSheet("CustomChart")
                .UseColumns("x", "y")
                .AsLineChart()
                .WithTitle("Test Chart")
                .WithCategoryAxisTitle("X Axis")
                .WithValueAxisTitle("Y Axis")
                .WithLegendPosition(ChartLegendPosition.Right))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        Assert.Equal("CustomChart", chartSheet.Name);
        
        var lineChart = (LineChart)chartSheet.Charts[0];
        Assert.Equal("Test Chart", lineChart.Title);
        Assert.Equal("X Axis", lineChart.CategoryAxisTitle);
        Assert.Equal("Y Axis", lineChart.ValueAxisTitle);
        Assert.Equal(ChartLegendPosition.Right, lineChart.LegendPosition);
    }

    [Fact]
    public void WithChart_AsScatterChartChained_ReturnsBuilderForChaining()
    {
        const string json = """[{"x": 1.5, "y": 2.5}, {"x": 2.5, "y": 3.5}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("x", "y")
                .AsScatterChart()
                .WithTitle("Scatter Plot")
                .WithCategoryAxisTitle("X Values")
                .WithValueAxisTitle("Y Values"))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var scatterChart = (ScatterChart)chartSheet.Charts[0];
        Assert.Equal("Scatter Plot", scatterChart.Title);
        Assert.Equal("X Values", scatterChart.CategoryAxisTitle);
        Assert.Equal("Y Values", scatterChart.ValueAxisTitle);
    }

    [Fact]
    public void WithChart_WithLegendPositionChained_ReturnsBuilderForChaining()
    {
        const string json = """[{"category": "A", "value": 10}, {"category": "B", "value": 20}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("category", "value")
                .AsPieChart()
                .WithLegendPosition(ChartLegendPosition.Left)
                .WithTitle("Distribution"))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var pieChart = (PieChart)chartSheet.Charts[0];
        Assert.Equal(ChartLegendPosition.Left, pieChart.LegendPosition);
        Assert.Equal("Distribution", pieChart.Title);
    }

    [Fact]
    public void WithFreezeHeaderRow_FreezesFirstRow()
    {
        const string json = """[{"name": "John", "age": 30}, {"name": "Jane", "age": 25}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithFreezeHeaderRow()
            .Build();
        
        var dataSheet = workbook.Sheets.First();
        Assert.NotNull(dataSheet.FrozenPane);
        Assert.Equal(1u, dataSheet.FrozenPane.Row);
        Assert.Equal(0u, dataSheet.FrozenPane.Column);
    }

    [Fact]
    public void WithChart_WithChartPosition_SetsCustomPosition()
    {
        const string json = """[{"x": 1, "y": 10}, {"x": 2, "y": 20}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("x", "y")
                .AsLineChart()
                .WithChartPosition(5, 5, 20, 25))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var lineChart = (LineChart)chartSheet.Charts[0];
        Assert.NotNull(lineChart.Position);
        Assert.Equal(5u, lineChart.Position.FromColumn);
        Assert.Equal(5u, lineChart.Position.FromRow);
        Assert.Equal(20u, lineChart.Position.ToColumn);
        Assert.Equal(25u, lineChart.Position.ToRow);
    }

    [Fact]
    public void WithChart_WithChartSize_SetsCustomSize()
    {
        const string json = """[{"a": 1, "b": 100}, {"a": 2, "b": 200}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("a", "b")
                .AsBarChart()
                .WithChartSize(800, 600))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var barChart = (BarChart)chartSheet.Charts[0];
        Assert.NotEqual(ChartSize.Default.WidthEmus, barChart.Size.WidthEmus);
        Assert.NotEqual(ChartSize.Default.HeightEmus, barChart.Size.HeightEmus);
    }

    [Fact]
    public void WithChart_WithDataLabels_ShowsLabels()
    {
        const string json = """[{"cat": "A", "val": 50}, {"cat": "B", "val": 75}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithChart(chart => chart
                .UseColumns("cat", "val")
                .AsPieChart()
                .WithDataLabels(true))
            .Build();
        
        var chartSheet = workbook.Sheets.Last();
        var pieChart = (PieChart)chartSheet.Charts[0];
        Assert.True(pieChart.ShowDataLabels);
    }

    [Fact]
    public void WithChart_AllPhase3Features_WorkTogether()
    {
        const string json = """[{"x": 1, "y": 100}, {"x": 2, "y": 200}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .WithFreezeHeaderRow()
            .WithChart(chart => chart
                .OnSheet("Analysis")
                .UseColumns("x", "y")
                .AsLineChart()
                .WithTitle("Data Analysis")
                .WithChartPosition(0, 2, 10, 15)
                .WithChartSize(600, 400)
                .WithDataLabels(true)
                .WithLegendPosition(ChartLegendPosition.Bottom))
            .Build();
        
        var dataSheet = workbook.Sheets.First();
        Assert.NotNull(dataSheet.FrozenPane);
        
        var chartSheet = workbook.Sheets.Last();
        var lineChart = (LineChart)chartSheet.Charts[0];
        Assert.Equal("Analysis", chartSheet.Name);
        Assert.NotNull(lineChart.Position);
        Assert.Equal(0u, lineChart.Position.FromColumn);
        Assert.Equal(2u, lineChart.Position.FromRow);
        Assert.NotEqual(ChartSize.Default, lineChart.Size);
        Assert.True(lineChart.ShowDataLabels);
    }
}

