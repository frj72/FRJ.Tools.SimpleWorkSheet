using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using SkiaSharp;

namespace FRJ.Tools.SimpleWorksheetTests;

public class TypefaceCacheTests
{
    [Fact]
    public void GetOrCreate_SameParameters_ReturnsSameInstance()
    {
        var typeface1 = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var typeface2 = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        Assert.Same(typeface1, typeface2);
    }

    [Fact]
    public void GetOrCreate_DifferentFamilyName_ReturnsDifferentInstance()
    {
        var typeface1 = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var typeface2 = TypefaceCache.GetOrCreate(
            "Calibri",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        Assert.NotSame(typeface1, typeface2);
    }

    [Fact]
    public void GetOrCreate_DifferentWeight_ReturnsDifferentInstance()
    {
        var typeface1 = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var typeface2 = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Bold,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        Assert.NotSame(typeface1, typeface2);
    }

    [Fact]
    public void GetOrCreate_DifferentSlant_ReturnsDifferentInstance()
    {
        var typeface1 = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var typeface2 = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Italic);

        Assert.NotSame(typeface1, typeface2);
    }

    [Fact]
    public void GetOrCreate_ReturnsValidTypeface()
    {
        var typeface = TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        Assert.NotNull(typeface);
        Assert.NotNull(typeface.FamilyName);
    }

    [Fact]
    public void GetOrCreate_WithDefaultFont_ReturnsValidTypeface()
    {
        var typeface = TypefaceCache.GetOrCreate(
            WorkSheetDefaults.FontName,
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        Assert.NotNull(typeface);
    }
}
