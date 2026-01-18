using System.Collections.Concurrent;
using SkiaSharp;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public static class TypefaceCache
{
    private record TypefaceKey(
        string FamilyName,
        SKFontStyleWeight Weight,
        SKFontStyleWidth Width,
        SKFontStyleSlant Slant);

    private static readonly ConcurrentDictionary<TypefaceKey, SKTypeface> Cache = new();
    private static readonly ReaderWriterLockSlim Lock = new(LockRecursionPolicy.NoRecursion);

    public static SKTypeface GetOrCreate(
        string familyName,
        SKFontStyleWeight weight,
        SKFontStyleWidth width,
        SKFontStyleSlant slant)
    {
        Lock.EnterReadLock();
        try
        {
            var key = new TypefaceKey(familyName, weight, width, slant);
            return Cache.GetOrAdd(key, k =>
                SKTypeface.FromFamilyName(k.FamilyName, k.Weight, k.Width, k.Slant));
        }
        finally
        {
            Lock.ExitReadLock();
        }
    }

    public static int GetCacheCount()
    {
        Lock.EnterReadLock();
        try
        {
            return Cache.Count;
        }
        finally
        {
            Lock.ExitReadLock();
        }
    }

    public static void ClearCache()
    {
        Lock.EnterWriteLock();
        try
        {
            foreach (var typeface in Cache.Values)
                typeface.Dispose();
            Cache.Clear();
        }
        finally
        {
            Lock.ExitWriteLock();
        }
    }
}
