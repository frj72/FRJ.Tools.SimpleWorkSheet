using Microsoft.Extensions.Caching.Memory;
using SkiaSharp;
using ZiggyCreatures.Caching.Fusion;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class TypefaceCache
{
    private static readonly FusionCache Cache = new(
        new FusionCacheOptions
        {
            CacheName = "TypefaceCache",
            DefaultEntryOptions = new FusionCacheEntryOptions
            {
                Duration = TimeSpan.MaxValue,
                Priority = CacheItemPriority.NeverRemove
            }
        });

    public static SKTypeface GetOrCreate(
        string familyName,
        SKFontStyleWeight weight,
        SKFontStyleWidth width,
        SKFontStyleSlant slant)
    {
        var key = $"{familyName}|{(int)weight}|{(int)width}|{(int)slant}";
        return Cache.GetOrSet(
            key,
            _ => SKTypeface.FromFamilyName(familyName, weight, width, slant));
    }
}
