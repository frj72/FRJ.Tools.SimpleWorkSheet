using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Charts;
using FRJ.Tools.SimpleWorkSheet.Components.Import;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.ImportExamples;

public class MultiTableMultiChartCommonSheetExample : IExample
{
    public string Name => "Multi Table Multi Chart Common Sheet";
    public string Description => "Creates 3 Generic Tables with business data, converts them to styled worksheets, and creates 4 charts on a single dashboard sheet";

    public void Run()
    {
        var salesTable = CreateSalesTable();
        var expensesTable = CreateExpensesTable();
        var budgetTable = CreateBudgetTable();

        var salesSheet = CreateStyledWorkSheet(salesTable, "SalesData");
        var expensesSheet = CreateStyledWorkSheet(expensesTable, "Expenses");
        var budgetSheet = CreateStyledWorkSheet(budgetTable, "Budget");

        var dashboardSheet = new WorkSheet("Dashboard");

        var barChart1 = BarChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 12),
                CellRange.FromBounds(4, 1, 4, 12))
            .WithDataSourceSheet("SalesData")
            .WithPosition(0, 0, 14, 18)
            .WithTitle("Monthly Sales Trend")
            .WithCategoryAxisTitle("Month")
            .WithValueAxisTitle("Revenue ($)")
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .WithMajorGridlines(true)
            .WithSeriesName("Total Sales");

        var barChart2 = BarChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 8),
                CellRange.FromBounds(1, 1, 1, 8))
            .WithDataSourceSheet("Expenses")
            .WithPosition(15, 0, 29, 18)
            .WithTitle("Expense Breakdown")
            .WithCategoryAxisTitle("Category")
            .WithValueAxisTitle("Amount ($)")
            .WithLegendPosition(ChartLegendPosition.Bottom)
            .WithMajorGridlines(true)
            .WithSeriesName("Expenses");

        var lineChart1 = LineChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 4),
                CellRange.FromBounds(3, 1, 3, 4))
            .WithDataSourceSheet("Budget")
            .WithPosition(0, 19, 14, 37)
            .WithTitle("Budget Variance Trend")
            .WithCategoryAxisTitle("Quarter")
            .WithValueAxisTitle("Variance (%)")
            .WithLegendPosition(ChartLegendPosition.Right)
            .WithMajorGridlines(true)
            .WithSeriesName("Budget vs Actual Variance");

        var lineChart2 = LineChart.Create()
            .WithDataRange(
                CellRange.FromBounds(0, 1, 0, 12),
                CellRange.FromBounds(3, 1, 3, 12))
            .WithDataSourceSheet("SalesData")
            .WithPosition(15, 19, 29, 37)
            .WithTitle("Units Sold Trend")
            .WithCategoryAxisTitle("Month")
            .WithValueAxisTitle("Units")
            .WithLegendPosition(ChartLegendPosition.Right)
            .WithMajorGridlines(true)
            .WithSeriesName("Units Sold");

        dashboardSheet.AddChart(barChart1);
        dashboardSheet.AddChart(barChart2);
        dashboardSheet.AddChart(lineChart1);
        dashboardSheet.AddChart(lineChart2);

        var workbook = new WorkBook("MultiTable Charts Demo", [salesSheet, expensesSheet, budgetSheet, dashboardSheet]);

        ExampleRunner.SaveWorkBook(workbook, "85_MultiTableMultiChartCommonSheet.xlsx");
    }

    private static GenericTable CreateSalesTable()
    {
        var table = GenericTable.Create("Month", "Region", "Product", "Units", "Revenue");
        table.AddRow("Jan", "North", "WidgetA", 125, 6250);
        table.AddRow("Feb", "North", "WidgetA", 132, 6600);
        table.AddRow("Mar", "North", "WidgetA", 148, 7400);
        table.AddRow("Apr", "North", "WidgetA", 156, 7800);
        table.AddRow("May", "North", "WidgetA", 174, 8700);
        table.AddRow("Jun", "North", "WidgetA", 189, 9450);
        table.AddRow("Jul", "North", "WidgetA", 203, 10150);
        table.AddRow("Aug", "North", "WidgetA", 218, 10900);
        table.AddRow("Sep", "North", "WidgetA", 234, 11700);
        table.AddRow("Oct", "North", "WidgetA", 247, 12350);
        table.AddRow("Nov", "North", "WidgetA", 261, 13050);
        table.AddRow("Dec", "North", "WidgetA", 278, 13900);
        return table;
    }

    private static GenericTable CreateExpensesTable()
    {
        var table = GenericTable.Create("Category", "Amount", "Department");
        table.AddRow("Salaries", 45000, "HR");
        table.AddRow("Marketing", 12000, "Sales");
        table.AddRow("Operations", 18500, "Operations");
        table.AddRow("IT", 9500, "IT");
        table.AddRow("Facilities", 7800, "Facilities");
        table.AddRow("Insurance", 3200, "Finance");
        table.AddRow("Training", 4100, "HR");
        table.AddRow("Miscellaneous", 2800, "Admin");
        return table;
    }

    private static GenericTable CreateBudgetTable()
    {
        var table = GenericTable.Create("Quarter", "Budget", "Actual", "VariancePct");
        table.AddRow("Q1", 150000, 145000, -3.33);
        table.AddRow("Q2", 160000, 168000, 5.0);
        table.AddRow("Q3", 175000, 172000, -1.71);
        table.AddRow("Q4", 180000, 185000, 2.78);
        return table;
    }

    private static WorkSheet CreateStyledWorkSheet(GenericTable table, string sheetName)
    {
        var sheet = new WorkSheet(sheetName);

        var headerBorders = CellBorders.Create(
            CellBorder.Create(Colors.White, CellBorderStyle.Thin),
            CellBorder.Create(Colors.White, CellBorderStyle.Thin),
            CellBorder.Create(Colors.White, CellBorderStyle.Thin),
            CellBorder.Create(Colors.White, CellBorderStyle.Thin));

        for (var col = 0; col < table.ColumnCount; col++)
        {
            var headerValue = new CellValue(table.GetHeader(col));
            sheet.AddCell(new CellPosition((uint)col, 0), headerValue, builder =>
                builder.WithStyle(style => style
                    .WithFillColor(Colors.UnitedNationsBlue)
                    .WithFont(font => font.Bold().WithColor(Colors.White))
                    .WithBorders(headerBorders)));
        }

        var dataBorders = CellBorders.Create(
            CellBorder.Create(Colors.AmericanSilver, CellBorderStyle.Thin),
            CellBorder.Create(Colors.AmericanSilver, CellBorderStyle.Thin),
            CellBorder.Create(Colors.AmericanSilver, CellBorderStyle.Thin),
            CellBorder.Create(Colors.AmericanSilver, CellBorderStyle.Thin));

        for (var row = 0; row < table.RowCount; row++)
        {
            var fillColor = row % 2 == 0 ? Colors.FreshAir : Colors.White;
            for (var col = 0; col < table.ColumnCount; col++)
            {
                var cellValue = table.GetValue(col, row);
                var value = cellValue ?? new CellValue("");
                sheet.AddCell(new CellPosition((uint)col, (uint)(row + 1)), value, builder =>
                    builder.WithStyle(style => style
                        .WithFillColor(fillColor)
                        .WithBorders(dataBorders)));
            }
        }

        return sheet;
    }
}
