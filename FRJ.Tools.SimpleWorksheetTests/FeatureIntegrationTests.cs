using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class FeatureIntegrationTests
{
    [Fact]
    public void NamedRangeWithFormula_Saves()
    {
        var sheet = new WorkSheet("Data");
        sheet.AddCell(new(0, 0), 100m, null);
        
        var workbook = new WorkBook("Test", [sheet]);
        workbook.AddNamedRange("MyValue", "Data", 0, 0, 0, 0);
        
        sheet.AddCell(new(1, 0), new CellFormula("=MyValue*2"), null);
        
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            Assert.True(File.Exists(tempPath));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void MultipleChartsOnSheet_Saves()
    {
        var sheet = new WorkSheet("Data");
        
        for (uint i = 0; i < 10; i++)
        {
            sheet.AddCell(new(0, i), $"Cat{i}", null);
            sheet.AddCell(new(1, i), i * 10, null);
        }
        
        var chart1 = BarChart.Create()
            .WithDataRange(CellRange.FromBounds(0, 0, 0, 9), CellRange.FromBounds(1, 0, 1, 9))
            .WithTitle("Chart 1")
            .WithPosition(4, 0, 10, 15);
        
        var chart2 = LineChart.Create()
            .WithDataRange(CellRange.FromBounds(0, 0, 0, 9), CellRange.FromBounds(1, 0, 1, 9))
            .WithTitle("Chart 2")
            .WithPosition(12, 0, 18, 15);
        
        sheet.AddChart(chart1);
        sheet.AddChart(chart2);
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            Assert.True(File.Exists(tempPath));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void CrossSheetFormula_RoundTrip()
    {
        var dataSheet = new WorkSheet("Data");
        var calcSheet = new WorkSheet("Calc");
        
        dataSheet.AddCell(new(0, 0), 100m, null);
        calcSheet.AddCell(new(0, 0), new CellFormula("=Data!A1*2"), null);
        
        var workbook = new WorkBook("Test", [dataSheet, calcSheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            var cell = reloaded.Sheets.ElementAt(1).Cells.Cells[new(0, 0)];
            Assert.True(cell.Value.Value.IsT5);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void MergedCellsWithHyperlink_BothPreserved()
    {
        var sheet = new WorkSheet("Links");
        
        sheet.AddCell(new(0, 0), "Link", cell => 
            cell.WithHyperlink("https://example.com", "Example"));
        sheet.MergeCells(CellRange.FromBounds(0, 0, 2, 0));
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            var cell = reloaded.Sheets.First().Cells.Cells[new(0, 0)];
            Assert.NotNull(cell.Hyperlink);
            Assert.Single(reloaded.Sheets.First().MergedCells);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void Validation_WithFormula_Saves()
    {
        var sheet = new WorkSheet("Data");
        var position = new CellPosition(0, 0);
        
        sheet.AddCell(position, new CellFormula("=SUM(A2:A6)"), null);
        
        var validation = CellValidation.WholeNumber(ValidationOperator.GreaterThan, 0);
        sheet.AddValidation(new CellRange(position, position), validation);
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            var cell = reloaded.Sheets.First().Cells.Cells[position];
            Assert.True(cell.Value.Value.IsT5);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void ValidationWithMergedCells_Saves()
    {
        var sheet = new WorkSheet("Overlap");
        
        sheet.AddCell(new(0, 0), "Merged", null);
        sheet.MergeCells(new CellRange(new CellPosition(0, 0), new CellPosition(2, 1)));
        
        var validation = CellValidation.TextLength(ValidationOperator.LessThan, 50);
        sheet.AddValidation(new CellRange(new CellPosition(0, 0), new CellPosition(2, 1)), validation);
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            Assert.True(File.Exists(tempPath));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void TableWithFrozenPanes_BothPreserved()
    {
        var sheet = new WorkSheet("Data");
        var range = CellRange.FromBounds(0, 0, 1, 5);
        
        sheet.AddTable("DataTable", range);
        sheet.FreezePanes(1, 0);
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.NotNull(reloaded.Sheets.First().FrozenPane);
            Assert.Single(reloaded.Sheets.First().Tables);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void MultipleValidations_OnDifferentRanges_Saves()
    {
        var sheet = new WorkSheet("Validations");
        
        var validation1 = CellValidation.WholeNumber(ValidationOperator.Between, 1, 100);
        sheet.AddValidation(CellRange.FromBounds(0, 0, 0, 10), validation1);
        
        var validation2 = CellValidation.List(["A", "B", "C"]);
        sheet.AddValidation(CellRange.FromBounds(1, 0, 1, 10), validation2);
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            Assert.True(File.Exists(tempPath));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void MultipleTablesOnSheet_AllPreserved()
    {
        var sheet = new WorkSheet("Tables");
        
        sheet.AddTable("Table1", CellRange.FromBounds(0, 0, 1, 5));
        sheet.AddTable("Table2", CellRange.FromBounds(4, 0, 5, 5));
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.Equal(2, reloaded.Sheets.First().Tables.Count);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void TabColorWithMultipleSheets_AllPreserved()
    {
        var sheet1 = new WorkSheet("Red");
        var sheet2 = new WorkSheet("Green");
        
        sheet1.SetTabColor("FF0000");
        sheet2.SetTabColor("00FF00");
        
        var workbook = new WorkBook("Test", [sheet1, sheet2]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.Equal("FF0000", reloaded.Sheets.ElementAt(0).TabColor);
            Assert.Equal("00FF00", reloaded.Sheets.ElementAt(1).TabColor);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void RowHeightsWithFrozenPanes_FrozenPreserved()
    {
        var sheet = new WorkSheet("Frozen");
        
        sheet.SetRowHeight(0, 30.0);
        sheet.SetRowHeight(1, 40.0);
        sheet.FreezePanes(2, 0);
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.NotNull(reloaded.Sheets.First().FrozenPane);
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void ColumnWidthsWithAutoFit_Preserved()
    {
        var sheet = new WorkSheet("Widths");
        
        sheet.AddCell(new(0, 0), "Test", null);
        sheet.SetColumnWidth(0, 25.0);
        sheet.AutoFitColumn(1);
        
        var workbook = new WorkBook("Test", [sheet]);
        var tempPath = Path.GetTempFileName() + ".xlsx";
        
        try
        {
            workbook.SaveToFile(tempPath);
            var reloaded = WorkBookReader.LoadFromFile(tempPath);
            
            Assert.True(reloaded.Sheets.First().ExplicitColumnWidths.ContainsKey(0));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }
}
