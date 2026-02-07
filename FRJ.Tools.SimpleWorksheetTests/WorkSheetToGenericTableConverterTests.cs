using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class WorkSheetToGenericTableConverterTests
{
    private static WorkSheet CreateTestWorkSheetWithHeaders()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.AddCell(new(0, 0), "Name", null);
        sheet.AddCell(new(1, 0), "Age", null);
        sheet.AddCell(new(2, 0), "Salary", null);

        sheet.AddCell(new(0, 1), "Alice", null);
        sheet.AddCell(new(1, 1), 28L, null);
        sheet.AddCell(new(2, 1), 75000m, null);

        sheet.AddCell(new(0, 2), "Bob", null);
        sheet.AddCell(new(1, 2), 35L, null);
        sheet.AddCell(new(2, 2), 65000m, null);

        return sheet;
    }

    private static WorkSheet CreateTestWorkSheetWithoutHeaders()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.AddCell(new(0, 0), 1000m, null);
        sheet.AddCell(new(1, 0), 2000m, null);
        sheet.AddCell(new(2, 0), 3000m, null);

        sheet.AddCell(new(0, 1), 1100m, null);
        sheet.AddCell(new(1, 1), 2100m, null);
        sheet.AddCell(new(2, 1), 3100m, null);

        return sheet;
    }

    [Fact]
    public void ConvertSheet_WithNoHeaders_GeneratesColumnNames()
    {
        var sheet = CreateTestWorkSheetWithoutHeaders();

        var table = WorkSheetToGenericTableConverter.Convert(sheet);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(2, table.RowCount);

        Assert.Equal("Column1", table.Headers[0]);
        Assert.Equal("Column2", table.Headers[1]);
        Assert.Equal("Column3", table.Headers[2]);
    }

    [Fact]
    public void ConvertSheet_WithFirstRowHeaders_ExtractsHeaders()
    {
        var sheet = CreateTestWorkSheetWithHeaders();

        var table = WorkSheetToGenericTableConverter.Convert(sheet, HeaderMode.FirstRow);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(2, table.RowCount);

        Assert.Equal("Name", table.Headers[0]);
        Assert.Equal("Age", table.Headers[1]);
        Assert.Equal("Salary", table.Headers[2]);
    }

    [Fact]
    public void ConvertSheet_WithAutoDetect_IdentifiesHeaders_WhenFirstRowIsAllStrings()
    {
        var sheet = CreateTestWorkSheetWithHeaders();

        var table = WorkSheetToGenericTableConverter.Convert(sheet, HeaderMode.AutoDetect);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(2, table.RowCount);

        Assert.Equal("Name", table.Headers[0]);
        Assert.Equal("Age", table.Headers[1]);
        Assert.Equal("Salary", table.Headers[2]);
    }

    [Fact]
    public void ConvertSheet_WithAutoDetect_FallsBackToGenerated_WhenFirstRowIsNotAllStrings()
    {
        var sheet = CreateTestWorkSheetWithoutHeaders();

        var table = WorkSheetToGenericTableConverter.Convert(sheet, HeaderMode.AutoDetect);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(2, table.RowCount);

        Assert.Equal("Column1", table.Headers[0]);
        Assert.Equal("Column2", table.Headers[1]);
        Assert.Equal("Column3", table.Headers[2]);
    }

    [Fact]
    public void ConvertSheet_PreservesStringValues()
    {
        var sheet = CreateTestWorkSheetWithHeaders();

        var table = WorkSheetToGenericTableConverter.Convert(sheet, HeaderMode.FirstRow);

        var name = table.GetValue("Name", 0);
        Assert.NotNull(name);
        Assert.True(name.IsString());
        Assert.Equal("Alice", name.AsString());
    }

    [Fact]
    public void ConvertSheet_PreservesLongValues()
    {
        var sheet = CreateTestWorkSheetWithHeaders();

        var table = WorkSheetToGenericTableConverter.Convert(sheet, HeaderMode.FirstRow);

        var age = table.GetValue("Age", 0);
        Assert.NotNull(age);
        Assert.True(age.IsLong());
        Assert.Equal(28L, age.AsLong());
    }

    [Fact]
    public void ConvertSheet_PreservesDecimalValues()
    {
        var sheet = CreateTestWorkSheetWithHeaders();

        var table = WorkSheetToGenericTableConverter.Convert(sheet, HeaderMode.FirstRow);

        var salary = table.GetValue("Salary", 0);
        Assert.NotNull(salary);
        Assert.True(salary.IsDecimal());
        Assert.Equal(75000m, salary.AsDecimal());
    }

    [Fact]
    public void ConvertSheet_PreservesDateTimeValues()
    {
        var sheet = new WorkSheet("TestSheet");
        var testDate = new DateTime(2024, 1, 15, 10, 30, 0);

        sheet.AddCell(new(0, 0), "Event", null);
        sheet.AddCell(new(1, 0), testDate, null);

        var table = WorkSheetToGenericTableConverter.Convert(sheet);

        var dateValue = table.GetValue(1, 0);
        Assert.NotNull(dateValue);
        Assert.True(dateValue.IsDateTime());
        Assert.Equal(testDate, dateValue.AsDateTime());
    }

    [Fact]
    public void ConvertSheet_ExtractsFormulaAsString()
    {
        var sheet = new WorkSheet("TestSheet");

        sheet.AddCell(new(0, 0), "Sum", null);
        sheet.AddCell(new(0, 1), new CellFormula("=SUM(A1:A2)"), null);

        var table = WorkSheetToGenericTableConverter.Convert(sheet);

        var formulaCell = table.GetValue(0, 1);
        Assert.NotNull(formulaCell);
        Assert.True(formulaCell.IsString());
        Assert.Equal("=SUM(A1:A2)", formulaCell.AsString());
    }

    [Fact]
    public void ConvertSheet_EmptySheet_ReturnsEmptyTable()
    {
        var sheet = new WorkSheet("EmptySheet");

        var table = WorkSheetToGenericTableConverter.Convert(sheet);

        Assert.Equal(0, table.ColumnCount);
        Assert.Equal(0, table.RowCount);
    }

    [Fact]
    public void ConvertSheet_SingleCell_CreatesTableWithOneCell()
    {
        var sheet = new WorkSheet("SingleCell");
        sheet.AddCell(new(0, 0), "OnlyValue", null);

        var table = WorkSheetToGenericTableConverter.Convert(sheet);

        Assert.Equal(1, table.ColumnCount);
        Assert.Equal(1, table.RowCount);

        var value = table.GetValue(0, 0);
        Assert.NotNull(value);
        Assert.Equal("OnlyValue", value.AsString());
    }

    [Fact]
    public void ConvertSheet_SparseData_InsertsNullsForMissingCells()
    {
        var sheet = new WorkSheet("Sparse");
        sheet.AddCell(new(0, 0), "A", null);
        sheet.AddCell(new(2, 2), "C", null);

        var table = WorkSheetToGenericTableConverter.Convert(sheet);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(3, table.RowCount);

        Assert.NotNull(table.GetValue(0, 0));
        Assert.Null(table.GetValue(1, 0));
        Assert.Null(table.GetValue(2, 0));

        Assert.Null(table.GetValue(0, 1));
        Assert.Null(table.GetValue(1, 1));
        Assert.Null(table.GetValue(2, 1));

        Assert.Null(table.GetValue(0, 2));
        Assert.Null(table.GetValue(1, 2));
        Assert.NotNull(table.GetValue(2, 2));
    }

    [Fact]
    public void ConvertWorkBook_MultipleSheets_ReturnsAllSheets()
    {
        var sheet1 = new WorkSheet("Sheet1");
        sheet1.AddCell(new(0, 0), "Data1", null);

        var sheet2 = new WorkSheet("Sheet2");
        sheet2.AddCell(new(0, 0), "Data2", null);

        var workbook = new WorkBook("TestWorkbook", [sheet1, sheet2]);

        var tables = WorkSheetToGenericTableConverter.ConvertWorkBook(workbook);

        Assert.Equal(2, tables.Count);
        Assert.True(tables.ContainsKey("Sheet1"));
        Assert.True(tables.ContainsKey("Sheet2"));
    }

    [Fact]
    public void ConvertWorkBook_SkipsEmptySheets()
    {
        var sheet1 = new WorkSheet("Sheet1");
        sheet1.AddCell(new(0, 0), "Data", null);

        var emptySheet = new WorkSheet("EmptySheet");

        var workbook = new WorkBook("TestWorkbook", [sheet1, emptySheet]);

        var tables = WorkSheetToGenericTableConverter.ConvertWorkBook(workbook);

        Assert.Single(tables);
        Assert.True(tables.ContainsKey("Sheet1"));
        Assert.False(tables.ContainsKey("EmptySheet"));
    }

    [Fact]
    public void ConvertWorkBook_PreservesSheetNames()
    {
        var sheet1 = new WorkSheet("Employees");
        sheet1.AddCell(new(0, 0), "Alice", null);

        var sheet2 = new WorkSheet("SalesData");
        sheet2.AddCell(new(0, 0), 1000m, null);

        var workbook = new WorkBook("TestWorkbook", [sheet1, sheet2]);

        var tables = WorkSheetToGenericTableConverter.ConvertWorkBook(workbook);

        Assert.Equal("Employees", tables.Keys.First());
        Assert.Equal("SalesData", tables.Keys.Skip(1).First());
    }

    [Fact]
    public void WorkBookReader_LoadAsGenericTables_LoadsFromFile()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Value", null);

        var workbook = new WorkBook("TestWorkbook", [sheet]);

        var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
        try
        {
            workbook.SaveToFile(tempPath);

            var tables = WorkBookReader.LoadAsGenericTables(tempPath);

            Assert.Single(tables);
            Assert.True(tables.ContainsKey("Test"));
            Assert.Equal(1, tables["Test"].RowCount);

            var value = tables["Test"].GetValue(0, 0);
            Assert.NotNull(value);
            Assert.Equal("Value", value.AsString());
        }
        finally
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }

    [Fact]
    public void WorkBookReader_LoadAsGenericTables_LoadsFromStream()
    {
        var sheet = new WorkSheet("Test");
        sheet.AddCell(new(0, 0), "Value", null);

        var workbook = new WorkBook("TestWorkbook", [sheet]);

        var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
        try
        {
            workbook.SaveToFile(tempPath);

            using var stream = File.OpenRead(tempPath);
            var tables = WorkBookReader.LoadAsGenericTables(stream);

            Assert.Single(tables);
            Assert.True(tables.ContainsKey("Test"));
        }
        finally
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }

    [Fact]
    public void ConvertSheet_WithDifferentHeaderModes_DataRowCountDiffers()
    {
        var sheet = CreateTestWorkSheetWithHeaders();

        var tableNone = WorkSheetToGenericTableConverter.Convert(sheet);
        var tableFirstRow = WorkSheetToGenericTableConverter.Convert(sheet, HeaderMode.FirstRow);

        Assert.Equal(3, tableNone.RowCount);
        Assert.Equal(2, tableFirstRow.RowCount);
    }

    [Fact]
    public void ConvertSheet_WithMergedCells_OnlyFirstCellHasValue()
    {
        var sheet = new WorkSheet("Merged");

        sheet.AddCell(new(0, 0), "MergedHeader", null);
        sheet.MergeCells(new(0, 0), new(1, 0));

        sheet.AddCell(new(0, 1), "Data1", null);
        sheet.AddCell(new(2, 1), "Data2", null);

        var table = WorkSheetToGenericTableConverter.Convert(sheet);

        Assert.Equal(3, table.ColumnCount);

        var mergedValue = table.GetValue(0, 0);
        Assert.NotNull(mergedValue);
        Assert.Equal("MergedHeader", mergedValue.AsString());

        Assert.Null(table.GetValue(1, 0));
    }
}
