using System.Runtime.InteropServices;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using SkiaSharp;

namespace FRJ.Tools.SimpleWorkSheet.DockerTest;

internal static class Program
{
    private static string GetOutputDirectory()
        => Directory.Exists("/output")
            ? "/output"
            : Path.Combine(Directory.GetCurrentDirectory(), "output");
    
    private static readonly string OutputDir = GetOutputDirectory();
    
    private static int Main()
    {
        try
        {
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine("FRJ.Tools.SimpleWorkSheet - Docker AutoFit Test");
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine();

            EnsureOutputDirectory();
            
            RunDiagnostics();
            
            GenerateWorkbooks();
            
            Console.WriteLine();
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine("Test completed successfully!");
            Console.WriteLine("=".PadRight(60, '='));
            
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine("FATAL ERROR:");
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }

    private static void EnsureOutputDirectory()
    {
        if (!Directory.Exists(OutputDir))
        {
            Directory.CreateDirectory(OutputDir);
            Console.WriteLine($"Created output directory: {OutputDir}");
        }
        else
            Console.WriteLine($"Using output directory: {OutputDir}");
        Console.WriteLine();
    }

    private static void RunDiagnostics()
    {
        Console.WriteLine("--- ENVIRONMENT DIAGNOSTICS ---");
        Console.WriteLine();
        
        Console.WriteLine($"Platform: {RuntimeInformation.OSDescription}");
        Console.WriteLine($"Architecture: {RuntimeInformation.OSArchitecture}");
        Console.WriteLine($"Framework: {RuntimeInformation.FrameworkDescription}");
        Console.WriteLine($"Runtime Identifier: {RuntimeInformation.RuntimeIdentifier}");
        Console.WriteLine();
        
        Console.WriteLine("--- SKIASHARP DIAGNOSTICS ---");
        Console.WriteLine();
        
        try
        {
            var skiaVersion = typeof(SKTypeface).Assembly.GetName().Version;
            Console.WriteLine($"SkiaSharp Version: {skiaVersion}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to get SkiaSharp version: {ex.Message}");
        }
        
        Console.WriteLine();
        Console.WriteLine("--- FONT LOADING TESTS ---");
        Console.WriteLine();
        
        TestFont("Aptos Narrow");
        TestFont("Arial");
        TestFont("Liberation Sans");
        TestFont("DejaVu Sans");
        TestFont("sans-serif");
        
        Console.WriteLine();
        Console.WriteLine("--- FONT MEASUREMENT TEST ---");
        Console.WriteLine();
        
        TestFontMeasurement();
        
        Console.WriteLine();
    }

    private static void TestFont(string fontName)
    {
        Console.Write($"Testing font '{fontName}'... ");
        
        try
        {
            using var typeface = SKTypeface.FromFamilyName(
                fontName,
                SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal,
                SKFontStyleSlant.Upright
            );
            
            if (typeface != null)
            {
                var actualFontFamily = typeface.FamilyName;
                Console.WriteLine(actualFontFamily.Equals(fontName, StringComparison.OrdinalIgnoreCase)
                    ? $"SUCCESS (loaded '{actualFontFamily}')"
                    : $"FALLBACK (requested '{fontName}', got '{actualFontFamily}')");
            }
            else
                Console.WriteLine("FAILED (typeface is null)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
        }
    }

    private static void TestFontMeasurement()
    {
        const string testText = "This is a simple calibration test";
        Console.WriteLine($"Measuring text: '{testText}'");
        Console.WriteLine();
        
        try
        {
            var width = EnvironmentSheetInfo.GetWidth("Aptos Narrow", 12, testText);
            Console.WriteLine($"Estimated width (Aptos Narrow, 12pt): {width:F2}");
            Console.WriteLine("Font measurement: SUCCESS");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Font measurement: FAILED");
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }

    private static void GenerateWorkbooks()
    {
        Console.WriteLine("--- GENERATING WORKBOOKS ---");
        Console.WriteLine();
        
        GenerateWorkbook("default", null, "calibration_default.xlsx");
        GenerateWorkbook("0.9", 0.9, "calibration_09.xlsx");
        GenerateWorkbook("1.1", 1.1, "calibration_11.xlsx");
        
        Console.WriteLine();
        Console.WriteLine("All workbooks generated successfully.");
    }

    private static void GenerateWorkbook(string description, double? calibrationFactor, string filename)
    {
        Console.Write($"Generating workbook with {description} calibration... ");
        
        try
        {
            var workbook = EnvironmentSheetInfo.CreateSimpleCalibrationWorkBook(calibrationFactor);
            
            var outputPath = Path.Combine(OutputDir, filename);
            workbook.SaveToFile(outputPath);
            
            var fileInfo = new FileInfo(outputPath);
            Console.WriteLine($"SUCCESS ({fileInfo.Length:N0} bytes)");
            Console.WriteLine($"  -> {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("FAILED");
            Console.WriteLine($"  Error: {ex.Message}");
            throw;
        }
    }
}
