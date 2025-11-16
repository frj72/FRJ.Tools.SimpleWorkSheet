using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class HyperlinkTests
{
    [Fact]
    public void WithHyperlink_SetsHyperlink()
    {
        var cell = CellBuilder.Create()
            .WithValue("Click here")
            .WithHyperlink("https://example.com")
            .Build();

        Assert.NotNull(cell.Hyperlink);
        Assert.Equal("https://example.com", cell.Hyperlink.Url);
        Assert.Null(cell.Hyperlink.Tooltip);
    }

    [Fact]
    public void WithHyperlink_WithTooltip_SetsTooltip()
    {
        var cell = CellBuilder.Create()
            .WithValue("Click here")
            .WithHyperlink("https://example.com", "Visit our site")
            .Build();

        Assert.NotNull(cell.Hyperlink);
        Assert.Equal("https://example.com", cell.Hyperlink.Url);
        Assert.Equal("Visit our site", cell.Hyperlink.Tooltip);
    }

    [Fact]
    public void CellWithHyperlink_CreatesValidExcelFile()
    {
        var sheet = new WorkSheet("Hyperlinks");

        sheet.AddCell(new(0, 0), "Visit Site", cell => cell
            .WithHyperlink("https://example.com", "Click to visit")
            .WithFont(font => font.WithColor("0000FF")));

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
        Assert.True(bytes.Length > 1000);
    }

    [Fact]
    public void MultipleHyperlinks_CreatesValidExcelFile()
    {
        var sheet = new WorkSheet("Hyperlinks");

        sheet.AddCell(new(0, 0), "Google", cell => cell
            .WithHyperlink("https://google.com"));

        sheet.AddCell(new(0, 1), "GitHub", cell => cell
            .WithHyperlink("https://github.com"));

        sheet.AddCell(new(0, 2), "Email", cell => cell
            .WithHyperlink("mailto:test@example.com"));

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void HyperlinkWithStyling_CreatesValidExcelFile()
    {
        var sheet = new WorkSheet("Hyperlinks");

        sheet.AddCell(new(0, 0), "Styled Link", cell => cell
            .WithHyperlink("https://example.com")
            .WithFont(font => font.WithColor("0000FF").Bold())
            .WithStyle(style => style.WithHorizontalAlignment(HorizontalAlignment.Center)));

        var workbook = new WorkBook("Test", [sheet]);
        var bytes = SheetConverter.ToBinaryExcelFile(workbook);

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void WithHyperlink_AllowsVeryLongUrls()
    {
        var longPath = new string('a', 2100);
        var url = $"https://example.com/{longPath}";

        var cell = CellBuilder.Create()
            .WithValue("Long URL")
            .WithHyperlink(url)
            .Build();

        Assert.Equal(url, cell.Hyperlink?.Url);
    }

    [Fact]
    public void WithHyperlink_AllowsSpecialCharacters()
    {
        const string url = "https://example.com/search?q=Résumé+😊";

        var cell = CellBuilder.Create()
            .WithValue("Special")
            .WithHyperlink(url)
            .Build();

        Assert.Equal(url, cell.Hyperlink?.Url);
    }

    [Fact]
    public void WithHyperlink_AllowsMailtoAddress()
    {
        const string url = "mailto:user@example.com";

        var cell = CellBuilder.Create()
            .WithValue("Email")
            .WithHyperlink(url)
            .Build();

        Assert.Equal(url, cell.Hyperlink?.Url);
    }

    [Fact]
    public void WithHyperlink_AllowsMailtoWithSubject()
    {
        const string url = "mailto:user@example.com?subject=Test";

        var cell = CellBuilder.Create()
            .WithValue("Email")
            .WithHyperlink(url, "Send email")
            .Build();

        Assert.Equal(url, cell.Hyperlink?.Url);
        Assert.Equal("Send email", cell.Hyperlink?.Tooltip);
    }

    [Fact]
    public void WithHyperlink_InvalidUrlString_IsStoredAsIs()
    {
        const string url = "not a valid url";

        var cell = CellBuilder.Create()
            .WithValue("Invalid")
            .WithHyperlink(url)
            .Build();

        Assert.Equal(url, cell.Hyperlink?.Url);
    }
}
