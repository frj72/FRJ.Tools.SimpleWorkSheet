using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorksheetTests;

public class FontFallbackConfigTests
{
    [Fact]
    public void GetDefaultFontName_ReturnsValidFont()
    {
        var fontName = FontFallbackConfig.GetDefaultFontName();

        Assert.NotNull(fontName);
        Assert.NotEmpty(fontName);
    }

    [Fact]
    public void GetDefaultFontName_ReturnsOneOfFallbackChain()
    {
        var fontName = FontFallbackConfig.GetDefaultFontName();
        var validFonts = new[] { "Aptos Narrow", "Calibri", "Arial" };

        Assert.Contains(fontName, validFonts);
    }

    [Fact]
    public void GetDefaultFontName_IsCached()
    {
        var fontName1 = FontFallbackConfig.GetDefaultFontName();
        var fontName2 = FontFallbackConfig.GetDefaultFontName();

        Assert.Same(fontName1, fontName2);
    }

    [Fact]
    public async Task GetDefaultFontName_ThreadSafe()
    {
        var results = new string[10];
        var tasks = new Task[10];

        for (var i = 0; i < 10; i++)
        {
            var index = i;
            tasks[i] = Task.Run(() =>
            {
                results[index] = FontFallbackConfig.GetDefaultFontName();
            });
        }

        await Task.WhenAll(tasks);

        var firstResult = results[0];
        Assert.All(results, r => Assert.Same(firstResult, r));
    }

    [Fact]
    public void Arial_IsAlwaysAvailable()
    {
        using var typeface = SkiaSharp.SKTypeface.FromFamilyName("Arial");

        Assert.NotNull(typeface);
    }
}
