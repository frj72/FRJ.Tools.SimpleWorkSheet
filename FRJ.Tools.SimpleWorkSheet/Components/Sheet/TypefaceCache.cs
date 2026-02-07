using SkiaSharp;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class TypefaceCache
{
    private record TypefaceKey(
        string FamilyName,
        SKFontStyleWeight Weight,
        SKFontStyleWidth Width,
        SKFontStyleSlant Slant);

    private static readonly Dictionary<TypefaceKey, Lazy<SKTypeface>> Cache = new();
    private static readonly Lock TypefaceCacheLock = new();

    public static SKTypeface GetOrCreate(
        string familyName,
        SKFontStyleWeight weight,
        SKFontStyleWidth width,
        SKFontStyleSlant slant)
    {
        lock (TypefaceCacheLock)
        {
            var key = new TypefaceKey(familyName, weight, width, slant);
            if (Cache.TryGetValue(key, out var lazy))
                return lazy.Value;

            lazy = new Lazy<SKTypeface>(
                () => SKTypeface.FromFamilyName(familyName, weight, width, slant));
            Cache[key] = lazy;
            return lazy.Value;
        }
    }

    public static int GetCacheCount()
    {
        lock (TypefaceCacheLock)
            return Cache.Count;
    }

    public static void ClearCache()
    {
        lock (TypefaceCacheLock)
        {
            foreach (var lazy in Cache.Values.Where(lazy => lazy.IsValueCreated))
                lazy.Value.Dispose();

            Cache.Clear();
        }
    }
}
