using SkiaSharp;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

internal static class FontFallbackConfig
{
    private static readonly string[] FallbackChain = ["Aptos Narrow", "Calibri", "Arial"];
    private static readonly Lazy<string> ResolvedDefaultFont = new(ResolveDefaultFont, true);

    public static string GetDefaultFontName() => ResolvedDefaultFont.Value;

    private static string ResolveDefaultFont()
    {
        foreach (var fontName in FallbackChain)
            if (IsFontAvailable(fontName))
                return fontName;

        return "Arial";
    }

    private static bool IsFontAvailable(string fontName)
    {
        using var typeface = SKTypeface.FromFamilyName(fontName);
        return typeface != null &&
               typeface.FamilyName.Equals(fontName, StringComparison.OrdinalIgnoreCase);
    }
}
