using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using SkiaSharp;

namespace FRJ.Tools.SimpleWorksheetTests;

[Collection("TypefaceCache")]
public class TypefaceCacheTests : IDisposable
{
    public TypefaceCacheTests()
    {
        TypefaceCache.ClearCache();
    }

    public void Dispose()
    {
        TypefaceCache.ClearCache();
        GC.SuppressFinalize(this);
    }

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
    public void GetCacheCount_InitiallyZero_ReturnsZero()
    {
        TypefaceCache.ClearCache();

        var count = TypefaceCache.GetCacheCount();

        Assert.Equal(0, count);
    }

    [Fact]
    public void GetCacheCount_AfterSingleCall_ReturnsOne()
    {
        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var count = TypefaceCache.GetCacheCount();

        Assert.Equal(1, count);
    }

    [Fact]
    public void GetCacheCount_AfterMultipleDifferentCalls_ReturnsCorrectCount()
    {
        TypefaceCache.ClearCache();

        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Bold,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        TypefaceCache.GetOrCreate(
            "Calibri",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var count = TypefaceCache.GetCacheCount();

        Assert.Equal(3, count);
    }

    [Fact]
    public void GetCacheCount_AfterDuplicateCalls_CountsOnlyUnique()
    {
        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var count = TypefaceCache.GetCacheCount();

        Assert.Equal(1, count);
    }

    [Fact]
    public void ClearCache_EmptiesCache()
    {
        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        TypefaceCache.GetOrCreate(
            "Calibri",
            SKFontStyleWeight.Bold,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Italic);

        TypefaceCache.ClearCache();

        var count = TypefaceCache.GetCacheCount();
        Assert.Equal(0, count);
    }

    [Fact]
    public void ClearCache_ThenGetOrCreate_RepopulatesCache()
    {
        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var countBefore = TypefaceCache.GetCacheCount();
        Assert.Equal(1, countBefore);

        TypefaceCache.ClearCache();

        var countAfterClear = TypefaceCache.GetCacheCount();
        Assert.Equal(0, countAfterClear);

        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        var countAfterRecreate = TypefaceCache.GetCacheCount();
        Assert.Equal(1, countAfterRecreate);
    }

    [Fact]
    public void GetOrCreate_ConcurrentCalls_ThreadSafe()
    {
        TypefaceCache.ClearCache();

        var results = new SKTypeface[100];

        Parallel.For(0, 100, i =>
        {
            results[i] = TypefaceCache.GetOrCreate(
                "Arial",
                SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal,
                SKFontStyleSlant.Upright);
        });

        var firstTypeface = results[0];
        Assert.All(results, typeface => Assert.Same(firstTypeface, typeface));

        var count = TypefaceCache.GetCacheCount();
        Assert.Equal(1, count);
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

    [Fact]
    public void ClearCache_CalledMultipleTimes_NoError()
    {
        TypefaceCache.GetOrCreate(
            "Arial",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        TypefaceCache.ClearCache();
        TypefaceCache.ClearCache();
        TypefaceCache.ClearCache();

        var count = TypefaceCache.GetCacheCount();
        Assert.Equal(0, count);
    }

    [Fact]
    public void ClearCache_OnEmptyCache_NoError()
    {
        TypefaceCache.ClearCache();

        var count = TypefaceCache.GetCacheCount();
        Assert.Equal(0, count);
    }

    [Fact]
    public async Task ThreadSafety_ConcurrentGetOrCreateAndClear_NoExceptions()
    {
        TypefaceCache.ClearCache();

        var exceptions = new System.Collections.Concurrent.ConcurrentBag<Exception>();
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        var testDuration = TimeSpan.FromSeconds(2);

        var tasks = new List<Task>();

        for (var i = 0; i < 4; i++)
            tasks.Add(Task.Run(() =>
            {
                try
                {
                    while (stopwatch.Elapsed < testDuration)
                    {
                        var typeface = TypefaceCache.GetOrCreate(
                            "Arial",
                            SKFontStyleWeight.Normal,
                            SKFontStyleWidth.Normal,
                            SKFontStyleSlant.Upright);
                        Assert.NotNull(typeface);
                    }
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                }
            }));

        tasks.Add(Task.Run(() =>
        {
            try
            {
                while (stopwatch.Elapsed < testDuration)
                {
                    TypefaceCache.ClearCache();
                    Thread.Sleep(50);
                }
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
        }));

        await Task.WhenAll(tasks.ToArray());

        Assert.Empty(exceptions);
    }

    [Fact]
    public void ThreadSafety_ConcurrentGetOrCreateMultipleFonts_AllSucceed()
    {
        TypefaceCache.ClearCache();

        var exceptions = new System.Collections.Concurrent.ConcurrentBag<Exception>();
        var fonts = new[] { "Arial", "Calibri", "Times New Roman", "Courier New" };
        var weights = new[] { SKFontStyleWeight.Normal, SKFontStyleWeight.Bold };
        var slants = new[] { SKFontStyleSlant.Upright, SKFontStyleSlant.Italic };

        Parallel.For(0, 1000, i =>
        {
            try
            {
                var font = fonts[i % fonts.Length];
                var weight = weights[i % weights.Length];
                var slant = slants[i % slants.Length];

                var typeface = TypefaceCache.GetOrCreate(font, weight, SKFontStyleWidth.Normal, slant);
                Assert.NotNull(typeface);
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
        });

        Assert.Empty(exceptions);
    }

    [Fact]
    public async Task ThreadSafety_GetCacheCountDuringConcurrentOperations_NoExceptions()
    {
        TypefaceCache.ClearCache();

        var exceptions = new System.Collections.Concurrent.ConcurrentBag<Exception>();
        var counts = new System.Collections.Concurrent.ConcurrentBag<int>();

        var tasks = new List<Task>();

        for (var i = 0; i < 3; i++)
            tasks.Add(Task.Run(() =>
            {
                try
                {
                    for (var j = 0; j < 100; j++)
                        TypefaceCache.GetOrCreate(
                            "Arial",
                            SKFontStyleWeight.Normal,
                            SKFontStyleWidth.Normal,
                            SKFontStyleSlant.Upright);
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                }
            }));

        tasks.Add(Task.Run(() =>
        {
            try
            {
                for (var j = 0; j < 100; j++)
                {
                    var count = TypefaceCache.GetCacheCount();
                    counts.Add(count);
                }
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
        }));

        await Task.WhenAll(tasks.ToArray());

        Assert.Empty(exceptions);
        Assert.NotEmpty(counts);
    }
}
