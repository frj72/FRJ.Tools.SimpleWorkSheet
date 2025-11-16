using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class TextAlignmentTests
{
    [Fact]
    public void WithHorizontalAlignment_SetsAlignment()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Center)
            .Build();

        Assert.Equal(HorizontalAlignment.Center, style.HorizontalAlignment);
    }

    [Fact]
    public void WithVerticalAlignment_SetsAlignment()
    {
        var style = CellStyleBuilder.Create()
            .WithVerticalAlignment(VerticalAlignment.Middle)
            .Build();

        Assert.Equal(VerticalAlignment.Middle, style.VerticalAlignment);
    }

    [Fact]
    public void WithAlignment_SetsBothAlignments()
    {
        var style = CellStyleBuilder.Create()
            .WithAlignment(HorizontalAlignment.Right, VerticalAlignment.Bottom)
            .Build();

        Assert.Equal(HorizontalAlignment.Right, style.HorizontalAlignment);
        Assert.Equal(VerticalAlignment.Bottom, style.VerticalAlignment);
    }

    [Fact]
    public void WithTextRotation_ValidRange_SetsRotation()
    {
        var style = CellStyleBuilder.Create()
            .WithTextRotation(45)
            .Build();

        Assert.Equal(45, style.TextRotation);
    }

    [Fact]
    public void WithTextRotation_NegativeValue_SetsRotation()
    {
        var style = CellStyleBuilder.Create()
            .WithTextRotation(-45)
            .Build();

        Assert.Equal(-45, style.TextRotation);
    }

    [Fact]
    public void WithTextRotation_OutOfRange_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithTextRotation(91));
        Assert.Throws<ArgumentException>(() => builder.WithTextRotation(-91));
    }

    [Fact]
    public void WithWrapText_True_EnablesWrapping()
    {
        var style = CellStyleBuilder.Create()
            .WithWrapText(true)
            .Build();

        Assert.True(style.WrapText);
    }

    [Fact]
    public void WithWrapText_False_DisablesWrapping()
    {
        var style = CellStyleBuilder.Create()
            .WithWrapText(false)
            .Build();

        Assert.False(style.WrapText);
    }

    [Fact]
    public void CellWithAlignment_CreatesValidExcelFile()
    {
        var sheet = new WorkSheet("Alignment");

        sheet.AddCell(new(0, 0), "Centered", cell => cell
            .WithStyle(style => style
                .WithHorizontalAlignment(HorizontalAlignment.Center)
                .WithVerticalAlignment(VerticalAlignment.Middle)));

        sheet.AddCell(new(1, 0), "Right Aligned", cell => cell
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Right)));

        sheet.AddCell(new(2, 0), "Wrapped Text", cell => cell
            .WithStyle(style => style.WithWrapText(true)));

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void CellWithTextRotation_CreatesValidExcelFile()
    {
        var sheet = new WorkSheet("Rotation");

        sheet.AddCell(new(0, 0), "Rotated 45°", cell => cell
            .WithStyle(style => style.WithTextRotation(45)));

        sheet.AddCell(new(1, 0), "Rotated -30°", cell => cell
            .WithStyle(style => style.WithTextRotation(-30)));

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void AllAlignmentOptions_CombinedStyle()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Center)
            .WithVerticalAlignment(VerticalAlignment.Middle)
            .WithTextRotation(15)
            .WithWrapText(true)
            .Build();

        Assert.Equal(HorizontalAlignment.Center, style.HorizontalAlignment);
        Assert.Equal(VerticalAlignment.Middle, style.VerticalAlignment);
        Assert.Equal(15, style.TextRotation);
        Assert.True(style.WrapText);
    }

    [Fact]
    public void HorizontalAlignment_Left_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Left)
            .Build();

        Assert.Equal(HorizontalAlignment.Left, style.HorizontalAlignment);
    }

    [Fact]
    public void HorizontalAlignment_Center_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Center)
            .Build();

        Assert.Equal(HorizontalAlignment.Center, style.HorizontalAlignment);
    }

    [Fact]
    public void HorizontalAlignment_Right_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Right)
            .Build();

        Assert.Equal(HorizontalAlignment.Right, style.HorizontalAlignment);
    }

    [Fact]
    public void HorizontalAlignment_Justify_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Justify)
            .Build();

        Assert.Equal(HorizontalAlignment.Justify, style.HorizontalAlignment);
    }

    [Fact]
    public void HorizontalAlignment_Fill_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Fill)
            .Build();

        Assert.Equal(HorizontalAlignment.Fill, style.HorizontalAlignment);
    }

    [Fact]
    public void HorizontalAlignment_Distributed_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithHorizontalAlignment(HorizontalAlignment.Distributed)
            .Build();

        Assert.Equal(HorizontalAlignment.Distributed, style.HorizontalAlignment);
    }

    [Fact]
    public void VerticalAlignment_Top_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithVerticalAlignment(VerticalAlignment.Top)
            .Build();

        Assert.Equal(VerticalAlignment.Top, style.VerticalAlignment);
    }

    [Fact]
    public void VerticalAlignment_Middle_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithVerticalAlignment(VerticalAlignment.Middle)
            .Build();

        Assert.Equal(VerticalAlignment.Middle, style.VerticalAlignment);
    }

    [Fact]
    public void VerticalAlignment_Bottom_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithVerticalAlignment(VerticalAlignment.Bottom)
            .Build();

        Assert.Equal(VerticalAlignment.Bottom, style.VerticalAlignment);
    }

    [Fact]
    public void VerticalAlignment_Justify_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithVerticalAlignment(VerticalAlignment.Justify)
            .Build();

        Assert.Equal(VerticalAlignment.Justify, style.VerticalAlignment);
    }

    [Fact]
    public void VerticalAlignment_Distributed_SetsCorrectly()
    {
        var style = CellStyleBuilder.Create()
            .WithVerticalAlignment(VerticalAlignment.Distributed)
            .Build();

        Assert.Equal(VerticalAlignment.Distributed, style.VerticalAlignment);
    }

    [Fact]
    public void TextRotation_Zero_NoRotation()
    {
        var style = CellStyleBuilder.Create()
            .WithTextRotation(0)
            .Build();

        Assert.Equal(0, style.TextRotation);
    }

    [Fact]
    public void TextRotation_MaxPositive_90Degrees()
    {
        var style = CellStyleBuilder.Create()
            .WithTextRotation(90)
            .Build();

        Assert.Equal(90, style.TextRotation);
    }

    [Fact]
    public void TextRotation_MaxNegative_Minus90Degrees()
    {
        var style = CellStyleBuilder.Create()
            .WithTextRotation(-90)
            .Build();

        Assert.Equal(-90, style.TextRotation);
    }

    [Fact]
    public void TextRotation_180Degrees_ThrowsException()
    {
        var builder = CellStyleBuilder.Create();

        Assert.Throws<ArgumentException>(() => builder.WithTextRotation(180));
    }

    [Fact]
    public void TextRotation_MidRangePositive_45Degrees()
    {
        var style = CellStyleBuilder.Create()
            .WithTextRotation(45)
            .Build();

        Assert.Equal(45, style.TextRotation);
    }

    [Fact]
    public void TextRotation_MidRangeNegative_Minus45Degrees()
    {
        var style = CellStyleBuilder.Create()
            .WithTextRotation(-45)
            .Build();

        Assert.Equal(-45, style.TextRotation);
    }
}
