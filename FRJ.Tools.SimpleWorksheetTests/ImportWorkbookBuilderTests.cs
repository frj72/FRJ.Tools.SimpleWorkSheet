using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

[Collection("TypefaceCache")]
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

    [Fact]
    public void FromCsv_WithHeader_CreatesWorkbook()
    {
        const string csv = "name,age\nJohn,30\nJane,25";
        
        var workbook = WorkbookBuilder.FromCsv(csv).Build();
        
        Assert.NotNull(workbook);
        Assert.Single(workbook.Sheets);
        var dataSheet = workbook.Sheets.First();
        Assert.Equal("name", dataSheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("John", dataSheet.GetValue(0, 1)?.Value.AsT2);
    }

    [Fact]
    public void FromCsv_WithChart_CreatesWorkbookWithChart()
    {
        const string csv = "month,sales\nJan,1000\nFeb,1500\nMar,1200";
        
        var workbook = WorkbookBuilder.FromCsv(csv)
            .WithChart(chart => chart
                .UseColumns("month", "sales")
                .AsLineChart()
                .WithTitle("Sales Trend"))
            .Build();
        
        Assert.Equal(2, workbook.Sheets.Count());
        var chartSheet = workbook.Sheets.Last();
        Assert.Single(chartSheet.Charts);
        var lineChart = (LineChart)chartSheet.Charts[0];
        Assert.Equal("Sales Trend", lineChart.Title);
    }

    [Fact]
    public void FromCsv_AllFeatures_WorkTogether()
    {
        const string csv = "date,amount,category\n2025-01-01,100,A\n2025-01-02,200,B";
        
        var workbook = WorkbookBuilder.FromCsv(csv)
            .WithWorkbookName("CSV Analysis")
            .WithDataSheetName("CSV Data")
            .WithFreezeHeaderRow()
            .WithHeaderStyle(style => style.WithFillColor("70AD47"))
            .AutoFitAllColumns()
            .WithChart(chart => chart
                .OnSheet("Trend")
                .UseColumns("date", "amount")
                .AsAreaChart()
                .WithTitle("Amount Trend"))
            .Build();
        
        Assert.Equal("CSV Analysis", workbook.Name);
        var dataSheet = workbook.Sheets.First();
        Assert.Equal("CSV Data", dataSheet.Name);
        Assert.NotNull(dataSheet.FrozenPane);
        
        var chartSheet = workbook.Sheets.Last();
        Assert.Equal("Trend", chartSheet.Name);
        Assert.Single(chartSheet.Charts);
    }

    [Fact]
    public void FromGenericTable_WithColumnOrder_OrdersColumnsCorrectly()
    {
        var table = GenericTable.Create("A", "B", "C");
        table.AddRow(new CellValue("1"), new CellValue("2"), new CellValue("3"));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithColumnOrder("C", "A", "B")
            .Build();
        
        var sheet = workbook.Sheets.First();
        Assert.Equal("C", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("A", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("B", sheet.GetValue(2, 0)?.Value.AsT2);
    }

    [Fact]
    public void FromGenericTable_WithExcludeColumns_ExcludesSpecifiedColumns()
    {
        var table = GenericTable.Create("Name", "Age", "Email", "Phone");
        table.AddRow(new CellValue("John"), new CellValue(30), new CellValue("john@test.com"), new CellValue("555-1234"));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithExcludeColumns("Email", "Phone")
            .Build();
        
        var sheet = workbook.Sheets.First();
        Assert.Equal("Name", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Age", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Null(sheet.GetValue(2, 0));
    }

    [Fact]
    public void FromGenericTable_WithIncludeColumns_IncludesOnlySpecifiedColumns()
    {
        var table = GenericTable.Create("ID", "Name", "Age", "Score");
        table.AddRow(new CellValue(1), new CellValue("Alice"), new CellValue(25), new CellValue(95.5m));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithIncludeColumns("Name", "Score")
            .Build();
        
        var sheet = workbook.Sheets.First();
        Assert.Equal("Name", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Score", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Null(sheet.GetValue(2, 0));
    }

    [Fact]
    public void FromGenericTable_WithDateFormat_FormatsDateColumns()
    {
        var date = new DateTime(2024, 11, 19, 14, 30, 0);
        var table = GenericTable.Create("Event", "Date");
        table.AddRow(new CellValue("Meeting"), new CellValue(date));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithDateFormat(DateFormat.IsoDate)
            .Build();
        
        var sheet = workbook.Sheets.First();
        var dateCell = sheet.Cells.Cells[new(1, 1)];
        Assert.NotNull(dateCell);
        Assert.Equal("yyyy-mm-dd", dateCell.Style.FormatCode);
    }

    [Fact]
    public void FromGenericTable_WithNumberFormat_FormatsSpecificColumn()
    {
        var table = GenericTable.Create("Item", "Price", "Quantity");
        table.AddRow(new CellValue("Widget"), new CellValue(19.99m), new CellValue(5));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithNumberFormat("Price", NumberFormat.Float2)
            .Build();
        
        var sheet = workbook.Sheets.First();
        var priceCell = sheet.Cells.Cells[new(1, 1)];
        Assert.NotNull(priceCell);
        Assert.Equal("0.00", priceCell.Style.FormatCode);
    }

    [Fact]
    public void FromGenericTable_WithConditionalStyle_AppliesStyleBasedOnCondition()
    {
        var table = GenericTable.Create("Student", "Grade");
        table.AddRow(new CellValue("Alice"), new CellValue(92m));
        table.AddRow(new CellValue("Bob"), new CellValue(67m));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithConditionalStyle("Grade",
                value => value.Value.AsT0 >= 90,
                style => style.WithFillColor(Colors.Green))
            .Build();
        
        var sheet = workbook.Sheets.First();
        var aliceGradeCell = sheet.Cells.Cells[new(1, 1)];
        var bobGradeCell = sheet.Cells.Cells[new(1, 2)];
        
        Assert.Equal(Colors.Green, aliceGradeCell.Style.FillColor);
        Assert.NotEqual(Colors.Green, bobGradeCell.Style.FillColor);
    }

    [Fact]
    public void FromGenericTable_AllNewMethods_WorkTogether()
    {
        var table = GenericTable.Create("Name", "Department", "Salary", "HireDate", "Status");
        table.AddRow(
            new CellValue("Alice"),
            new CellValue("Engineering"),
            new CellValue(75000m),
            new CellValue(new DateTime(2020, 1, 15)),
            new CellValue("Active"));
        table.AddRow(
            new CellValue("Bob"),
            new CellValue("Sales"),
            new CellValue(65000m),
            new CellValue(new DateTime(2019, 6, 1)),
            new CellValue("Active"));
        
        var workbook = WorkbookBuilder.FromGenericTable(table)
            .WithWorkbookName("Employee Report")
            .WithDataSheetName("Employees")
            .WithColumnOrder("Name", "Salary", "HireDate", "Department")
            .WithExcludeColumns("Status")
            .WithDateFormat(DateFormat.IsoDate)
            .WithNumberFormat("Salary", NumberFormat.Float2)
            .WithConditionalStyle("Salary",
                value => value.Value.AsT0 > 70000,
                style => style.WithFillColor(Colors.Yellow))
            .WithFreezeHeaderRow()
            .AutoFitAllColumns()
            .Build();
        
        Assert.Equal("Employee Report", workbook.Name);
        var sheet = workbook.Sheets.First();
        Assert.Equal("Employees", sheet.Name);
        
        Assert.Equal("Name", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Salary", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("HireDate", sheet.GetValue(2, 0)?.Value.AsT2);
        Assert.Equal("Department", sheet.GetValue(3, 0)?.Value.AsT2);
        
        Assert.NotNull(sheet.FrozenPane);
        Assert.Equal(1u, sheet.FrozenPane.Row);
        
        var salaryCell = sheet.Cells.Cells[new(1, 1)];
        Assert.Equal("0.00", salaryCell.Style.FormatCode);
        Assert.Equal(Colors.Yellow, salaryCell.Style.FillColor);
        
        var dateCell = sheet.Cells.Cells[new(2, 1)];
        Assert.Equal("yyyy-mm-dd", dateCell.Style.FormatCode);
    }

    [Fact]
    public void FromClass_SimpleObject_CreatesWorkbook()
    {
        var person = new { Name = "John", Age = 30, Email = "john@test.com" };
        
        var workbook = WorkbookBuilder.FromClass(person).Build();
        
        Assert.NotNull(workbook);
        Assert.Single(workbook.Sheets);
    }

    [Fact]
    public void FromClass_ListOfObjects_CreatesRows()
    {
        var data = new[]
        {
            new TestProduct { Id = 1, Name = "First" },
            new TestProduct { Id = 2, Name = "Second" }
        };
        
        var workbook = WorkbookBuilder.FromClass(data).Build();
        
        var sheet = workbook.Sheets.First();
        Assert.NotNull(sheet.GetValue(0, 1));
        Assert.NotNull(sheet.GetValue(0, 2));
    }

    [Fact]
    public void FromClass_WithConfiguration_AppliesSettings()
    {
        var products = new TestProductCatalog
        {
            Products =
            [
                new TestProduct { Name = "Product A", Price = 19.99m },
                new TestProduct { Name = "Product B", Price = 29.99m }
            ]
        };
        
        var workbook = WorkbookBuilder.FromClass(products)
            .WithWorkbookName("Product List")
            .WithDataSheetName("Products")
            .Build();
        
        Assert.Equal("Product List", workbook.Name);
        Assert.Equal("Products", workbook.Sheets.First().Name);
    }

    [Fact]
    public void FromClass_WithParsers_TransformsData()
    {
        var sales = new[]
        {
            new TestProduct { Name = "Widget", Price = 100m },
            new TestProduct { Name = "Gadget", Price = 200m }
        };
        
        var workbook = WorkbookBuilder.FromClass(sales)
            .WithColumnParser("Price", value => new(value.Value.AsT0 * 1.1m))
            .Build();
        
        var sheet = workbook.Sheets.First();
        var amountValue = sheet.GetValue(2, 1);
        Assert.NotNull(amountValue);
        Assert.Equal(110m, amountValue.Value.AsT0);
    }

    [Fact]
    public void FromClass_NestedObject_FlattensCorrectly()
    {
        var employees = new[]
        {
            new TestEmployeeWithDetails { Name = "Alice", Details = new TestEmployeeDetails { Age = 30, Department = "IT" } }
        };
        
        var workbook = WorkbookBuilder.FromClass(employees).Build();
        
        var sheet = workbook.Sheets.First();
        var detailsAgeHeader = sheet.GetValue(1, 0);
        Assert.NotNull(detailsAgeHeader);
        Assert.Contains("Age", detailsAgeHeader.Value.AsT2);
    }

    [Fact]
    public void FromClass_WithChart_CreatesChartSheet()
    {
        var monthlyData = new[]
        {
            new TestMonthlyData { Month = "Jan", Sales = 1000 },
            new TestMonthlyData { Month = "Feb", Sales = 1500 }
        };
        
        var workbook = WorkbookBuilder.FromClass(monthlyData)
            .WithChart(chart => chart
                .UseColumns("Month", "Sales")
                .AsLineChart()
                .WithTitle("Monthly Sales"))
            .Build();
        
        Assert.Equal(2, workbook.Sheets.Count());
        var chartSheet = workbook.Sheets.Last();
        Assert.Single(chartSheet.Charts);
    }

    [Fact]
    public void FromClass_AllFeatures_WorkTogether()
    {
        var catalog = new TestProductCatalog
        {
            Products =
            [
                new TestProduct { Id = 1, Name = "Laptop", Price = 999.99m, InStock = true },
                new TestProduct { Id = 2, Name = "Mouse", Price = 24.99m, InStock = false }
            ]
        };
        
        var workbook = WorkbookBuilder.FromClass(catalog)
            .WithWorkbookName("Catalog")
            .WithDataSheetName("Items")
            .WithHeaderStyle(style => style.WithFillColor("4472C4"))
            .WithNumberFormat("Products.Price", NumberFormat.Float2)
            .WithFreezeHeaderRow()
            .AutoFitAllColumns()
            .Build();
        
        Assert.Equal("Catalog", workbook.Name);
        var sheet = workbook.Sheets.First();
        Assert.Equal("Items", sheet.Name);
        Assert.NotNull(sheet.FrozenPane);
        Assert.NotEmpty(sheet.ExplicitColumnWidths);
    }

    [Fact]
    public void FromClass_EmptyObject_CreatesEmptyWorkbook()
    {
        var emptyObj = new { };
        
        var workbook = WorkbookBuilder.FromClass(emptyObj).Build();
        
        Assert.NotNull(workbook);
        Assert.Single(workbook.Sheets);
    }

    [Fact]
    public void FromClass_ProductCatalogWithProperties_SerializesAllProperties()
    {
        var products = new[]
        {
            new TestProduct
            {
                Id = 1,
                Name = "Laptop",
                Price = 999.99m,
                InStock = true,
                Category = "Electronics"
            }
        };

        var workbook = WorkbookBuilder.FromClass(products).Build();
        var sheet = workbook.Sheets.First();

        Assert.Equal("Id", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Name", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("Price", sheet.GetValue(2, 0)?.Value.AsT2);
        Assert.Equal("InStock", sheet.GetValue(3, 0)?.Value.AsT2);
        Assert.Equal("Category", sheet.GetValue(4, 0)?.Value.AsT2);

        Assert.Equal(1m, sheet.GetValue(0, 1)?.Value.AsT0);
        Assert.Equal("Laptop", sheet.GetValue(1, 1)?.Value.AsT2);
        Assert.Equal(999.99m, sheet.GetValue(2, 1)?.Value.AsT0);
        Assert.Equal("TRUE", sheet.GetValue(3, 1)?.Value.AsT2);
        Assert.Equal("Electronics", sheet.GetValue(4, 1)?.Value.AsT2);
    }

    [Fact]
    public void FromClass_ProductCatalogGetterUsage_AllPropertiesReadCorrectly()
    {
        var product = new TestProduct
        {
            Id = 42,
            Name = "Test Product",
            Price = 123.45m,
            InStock = false,
            Category = "Test Category"
        };

        Assert.Equal(42, product.Id);
        Assert.Equal("Test Product", product.Name);
        Assert.Equal(123.45m, product.Price);
        Assert.False(product.InStock);
        Assert.Equal("Test Category", product.Category);

        var catalog = new TestProductCatalog { Products = [product] };
        Assert.Single(catalog.Products);
        Assert.Equal(42, catalog.Products[0].Id);
    }

    private sealed class TestProductCatalog
    {
        public List<TestProduct> Products { get; init; } = [];
    }

    private sealed class TestProduct
    {
        public int Id { get; init; }
        public string Name { get; init; } = string.Empty;
        public decimal Price { get; init; }
        public bool InStock { get; init; }
        public string Category { get; init; } = string.Empty;
    }

    private sealed class TestEmployeeWithDetails
    {
        public string Name { get; init; } = string.Empty;
        public TestEmployeeDetails Details { get; init; } = new();
    }

    private sealed class TestEmployeeDetails
    {
        public int Age { get; init; }
        public string Department { get; init; } = string.Empty;
    }

    private sealed class TestMonthlyData
    {
        public string Month { get; init; } = string.Empty;
        public int Sales { get; init; }
    }

    [Fact]
    public void TestEmployeeWithDetails_PropertiesWork()
    {
        var details = new TestEmployeeDetails { Age = 25, Department = "IT" };
        var employee = new TestEmployeeWithDetails { Name = "Bob", Details = details };
        
        Assert.Equal("Bob", employee.Name);
        Assert.Equal(25, employee.Details.Age);
        Assert.Equal("IT", employee.Details.Department);
    }

    [Fact]
    public void TestMonthlyData_PropertiesWork()
    {
        var data = new TestMonthlyData { Month = "March", Sales = 5000 };
        
        Assert.Equal("March", data.Month);
        Assert.Equal(5000, data.Sales);
    }

    [Fact]
    public void AutoFitAllColumns_WithCalibration_EnablesAutoFit()
    {
        const string json = """[{"name": "John", "age": 30, "city": "New York"}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .AutoFitAllColumns(0.9)
            .Build();
        
        Assert.NotNull(workbook);
        var sheet = workbook.Sheets.First();
        Assert.NotEmpty(sheet.ExplicitColumnWidths);
    }

    [Fact]
    public void AutoFitAllColumns_WithCalibrationAndBaseLine_AppliesBoth()
    {
        const string json = """[{"product": "Laptop", "price": 999.99, "quantity": 5}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .AutoFitAllColumns(1.2, 8.0)
            .Build();
        
        Assert.NotNull(workbook);
        var sheet = workbook.Sheets.First();
        Assert.NotEmpty(sheet.ExplicitColumnWidths);
        
        foreach (var width in sheet.ExplicitColumnWidths.Values)
        {
            Assert.True(width.IsT0);
            Assert.True(width.AsT0 > 8.0);
        }
    }

    [Fact]
    public void AutoFitAllColumns_WithNegativeBaseLine_AppliesCorrectly()
    {
        const string json = """[{"name": "Test", "value": 100}]""";
        
        var workbook = WorkbookBuilder.FromJson(json)
            .AutoFitAllColumns(1.0, -3.0)
            .Build();
        
        Assert.NotNull(workbook);
        var sheet = workbook.Sheets.First();
        Assert.NotEmpty(sheet.ExplicitColumnWidths);
        
        foreach (var width in sheet.ExplicitColumnWidths.Values)
        {
            Assert.True(width.IsT0);
            Assert.True(width.AsT0 > 0);
        }
    }

    [Fact]
    public void AutoFitAllColumns_WithZeroBaseLine_EquivalentToCalibrationOnly()
    {
        const string json = """[{"test": "value"}]""";
        
        var workbook1 = WorkbookBuilder.FromJson(json)
            .AutoFitAllColumns(1.5)
            .Build();
        
        var workbook2 = WorkbookBuilder.FromJson(json)
            .AutoFitAllColumns(1.5, 0.0)
            .Build();
        
        var width1 = workbook1.Sheets.First().ExplicitColumnWidths[0].AsT0;
        var width2 = workbook2.Sheets.First().ExplicitColumnWidths[0].AsT0;
        
        Assert.Equal(width1, width2, 0.001);
    }
}



