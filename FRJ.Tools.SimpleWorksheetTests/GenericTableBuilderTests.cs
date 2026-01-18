using FRJ.Tools.SimpleWorkSheet.Components.Formatting;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

[Collection("TypefaceCache")]
public class GenericTableBuilderTests
{
    [Fact]
    public void FromGenericTable_EmptyTable_ReturnsEmptySheet()
    {
        var table = GenericTable.Create();
        
        var sheet = GenericTableBuilder.FromGenericTable(table).Build();
        
        Assert.NotNull(sheet);
        Assert.Empty(sheet.Cells.Cells);
    }

    [Fact]
    public void FromGenericTable_WithData_CreatesSheet()
    {
        var table = GenericTable.Create("Name", "Age", "City");
        table.AddRow(new CellValue("John"), new CellValue(30), new CellValue("NYC"));
        table.AddRow(new CellValue("Jane"), new CellValue(25), new CellValue("LA"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table).Build();
        
        Assert.Equal("Name", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Age", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("City", sheet.GetValue(2, 0)?.Value.AsT2);
        Assert.Equal("John", sheet.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal(30, sheet.GetValue(1, 1)?.Value.AsT1);
        Assert.Equal("NYC", sheet.GetValue(2, 1)?.Value.AsT2);
        Assert.Equal("Jane", sheet.GetValue(0, 2)?.Value.AsT2);
        Assert.Equal(25, sheet.GetValue(1, 2)?.Value.AsT1);
        Assert.Equal("LA", sheet.GetValue(2, 2)?.Value.AsT2);
    }

    [Fact]
    public void WithSheetName_ValidName_SetsSheetName()
    {
        var table = GenericTable.Create("A");
        table.AddRow(new CellValue("X"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithSheetName("TestSheet")
            .Build();
        
        Assert.Equal("TestSheet", sheet.Name);
    }

    [Fact]
    public void WithSheetName_EmptyName_ThrowsArgumentException()
    {
        var table = GenericTable.Create("A");
        var builder = GenericTableBuilder.FromGenericTable(table);
        
        Assert.Throws<ArgumentException>(() => builder.WithSheetName(""));
    }

    [Fact]
    public void WithHeaderStyle_AppliesStyleToHeaders()
    {
        var table = GenericTable.Create("Name", "Age");
        table.AddRow(new CellValue("John"), new CellValue(30));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithHeaderStyle(style => style.WithFont(font => font.Bold()))
            .Build();
        
        var headerCell = sheet.Cells.Cells[new(0, 0)];
        Assert.NotNull(headerCell);
        Assert.NotNull(headerCell.Font);
        Assert.True(headerCell.Font.Bold);
    }

    [Fact]
    public void WithColumnParser_ParsesColumnValues()
    {
        var table = GenericTable.Create("Name", "Salary");
        table.AddRow(new CellValue("John"), new CellValue(1000m));
        table.AddRow(new CellValue("Jane"), new CellValue(2000m));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithColumnParser("Salary", value => new(value.Value.AsT0 * 1.1m))
            .Build();
        
        Assert.Equal(1100m, sheet.GetValue(1, 1)?.Value.AsT0);
        Assert.Equal(2200m, sheet.GetValue(1, 2)?.Value.AsT0);
    }

    [Fact]
    public void AutoFitAllColumns_EnablesAutoFit()
    {
        var table = GenericTable.Create("A", "B");
        table.AddRow(new CellValue("Short"), new CellValue("VeryLongValue"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .AutoFitAllColumns()
            .Build();
        
        Assert.NotNull(sheet);
        Assert.Equal("A", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("B", sheet.GetValue(1, 0)?.Value.AsT2);
    }

    [Fact]
    public void AutoFitAllColumns_WithCalibration_EnablesAutoFit()
    {
        var table = GenericTable.Create("A", "B");
        table.AddRow(new CellValue("Short"), new CellValue("VeryLongValue"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .AutoFitAllColumns(0.9)
            .Build();
        
        Assert.NotNull(sheet);
        Assert.Equal("A", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("B", sheet.GetValue(1, 0)?.Value.AsT2);
    }

    [Fact]
    public void WithColumnOrder_ReordersColumns()
    {
        var table = GenericTable.Create("A", "B", "C");
        table.AddRow(new CellValue("1"), new CellValue("2"), new CellValue("3"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithColumnOrder("C", "A", "B")
            .Build();
        
        Assert.Equal("C", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("A", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("B", sheet.GetValue(2, 0)?.Value.AsT2);
        Assert.Equal("3", sheet.GetValue(0, 1)?.Value.AsT2);
        Assert.Equal("1", sheet.GetValue(1, 1)?.Value.AsT2);
        Assert.Equal("2", sheet.GetValue(2, 1)?.Value.AsT2);
    }

    [Fact]
    public void WithExcludeColumns_ExcludesSpecifiedColumns()
    {
        var table = GenericTable.Create("A", "B", "C");
        table.AddRow(new CellValue("1"), new CellValue("2"), new CellValue("3"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithExcludeColumns("B")
            .Build();
        
        Assert.Equal("A", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("C", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Null(sheet.GetValue(2, 0));
    }

    [Fact]
    public void WithIncludeColumns_IncludesOnlySpecifiedColumns()
    {
        var table = GenericTable.Create("A", "B", "C", "D");
        table.AddRow(new CellValue("1"), new CellValue("2"), new CellValue("3"), new CellValue("4"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithIncludeColumns("A", "C")
            .Build();
        
        Assert.Equal("A", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("C", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Null(sheet.GetValue(2, 0));
    }

    [Fact]
    public void WithDateFormat_FormatsDateColumns()
    {
        var date = new DateTime(2024, 1, 15, 10, 30, 0);
        var table = GenericTable.Create("Name", "Date");
        table.AddRow(new CellValue("Event"), new CellValue(date));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithDateFormat(DateFormat.IsoDate)
            .Build();
        
        var dateCell = sheet.Cells.Cells[new(1, 1)];
        Assert.NotNull(dateCell);
        Assert.Equal("yyyy-mm-dd", dateCell.Style.FormatCode);
    }

    [Fact]
    public void WithNumberFormat_FormatsNumberColumn()
    {
        var table = GenericTable.Create("Name", "Price");
        table.AddRow(new CellValue("Item1"), new CellValue(99.99m));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithNumberFormat("Price", NumberFormat.Float2)
            .Build();
        
        var priceCell = sheet.Cells.Cells[new(1, 1)];
        Assert.NotNull(priceCell);
        Assert.Equal("0.00", priceCell.Style.FormatCode);
    }

    [Fact]
    public void WithConditionalStyle_AppliesStyleBasedOnCondition()
    {
        var table = GenericTable.Create("Name", "Score");
        table.AddRow(new CellValue("Test1"), new CellValue(85m));
        table.AddRow(new CellValue("Test2"), new CellValue(45m));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithConditionalStyle("Score", 
                value => value.Value.AsT0 >= 50, 
                style => style.WithFillColor(Colors.Green))
            .Build();
        
        var cell1 = sheet.Cells.Cells[new(1, 1)];
        var cell2 = sheet.Cells.Cells[new(1, 2)];
        
        Assert.NotNull(cell1);
        Assert.Equal(Colors.Green, cell1.Style.FillColor);
        Assert.NotNull(cell2);
        Assert.NotEqual(Colors.Green, cell2.Style.FillColor);
    }

    [Fact]
    public void WithPreserveOriginalValue_PreservesOriginalValues()
    {
        var table = GenericTable.Create("Name", "Value");
        table.AddRow(new CellValue("Item"), new CellValue(123.456m));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithPreserveOriginalValue(true)
            .Build();
        
        var cell = sheet.Cells.Cells[new(1, 1)];
        Assert.NotNull(cell);
        Assert.NotNull(cell.Metadata);
        Assert.NotNull(cell.Metadata.OriginalValue);
    }

    [Fact]
    public void WithPreserveOriginalValue_False_DoesNotPreserveOriginalValues()
    {
        var table = GenericTable.Create("Name", "Value");
        table.AddRow(new CellValue("Item"), new CellValue(123.456m));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithPreserveOriginalValue(false)
            .Build();
        
        var cell = sheet.Cells.Cells[new(1, 1)];
        Assert.NotNull(cell);
        Assert.Null(cell.Metadata);
    }

    [Fact]
    public void WithNullValues_HandlesNullCells()
    {
        var table = GenericTable.Create("A", "B", "C");
        table.AddRow(new CellValue("X"), null, new CellValue("Z"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table).Build();
        
        Assert.Equal("X", sheet.GetValue(0, 1)?.Value.AsT2);
        Assert.Null(sheet.GetValue(1, 1));
        Assert.Equal("Z", sheet.GetValue(2, 1)?.Value.AsT2);
    }

    [Fact]
    public void ChainedOperations_AllFeaturesWork()
    {
        var table = GenericTable.Create("Name", "Age", "Salary", "Department");
        table.AddRow(new CellValue("John"), new CellValue(30), new CellValue(50000m), new CellValue("IT"));
        table.AddRow(new CellValue("Jane"), new CellValue(25), new CellValue(60000m), new CellValue("HR"));
        
        var sheet = GenericTableBuilder.FromGenericTable(table)
            .WithSheetName("Employees")
            .WithHeaderStyle(style => style.WithFont(font => font.Bold()).WithFillColor(Colors.Yellow))
            .WithColumnParser("Salary", value => new(value.Value.AsT0 * 1.05m))
            .WithNumberFormat("Salary", NumberFormat.Float2)
            .WithExcludeColumns("Department")
            .WithColumnOrder("Name", "Salary", "Age")
            .WithConditionalStyle("Salary", 
                value => value.Value.AsT0 > 55000, 
                style => style.WithFillColor(Colors.Green))
            .AutoFitAllColumns()
            .Build();
        
        Assert.Equal("Employees", sheet.Name);
        Assert.Equal("Name", sheet.GetValue(0, 0)?.Value.AsT2);
        Assert.Equal("Salary", sheet.GetValue(1, 0)?.Value.AsT2);
        Assert.Equal("Age", sheet.GetValue(2, 0)?.Value.AsT2);
        Assert.Equal(52500m, sheet.GetValue(1, 1)?.Value.AsT0);
        Assert.Equal(63000m, sheet.GetValue(1, 2)?.Value.AsT0);
    }
}
