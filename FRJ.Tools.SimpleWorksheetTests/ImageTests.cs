using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.LowLevel;

namespace FRJ.Tools.SimpleWorksheetTests;

public class ImageTests
{
    private static byte[] CreateTestImageData() => [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    [Fact]
    public void AddImage_ValidImage_AddsToSheet()
    {
        var sheet = new WorkSheet("TestSheet");
        var imageData = CreateTestImageData();
        var image = new WorksheetImage(imageData, ImageFormat.Png, new(0, 0), 100, 100);
        
        sheet.AddImage(image);
        
        Assert.Single(sheet.Images);
        Assert.Equal(image, sheet.Images[0]);
    }

    [Fact]
    public void AddImage_FromFile_AddsImage()
    {
        var sheet = new WorkSheet("TestSheet");
        var tempFile = Path.GetTempFileName() + ".png";
        
        try
        {
            File.WriteAllBytes(tempFile, CreateTestImageData());
            
            sheet.AddImage(tempFile, new(0, 0), 150, 150);
            
            Assert.Single(sheet.Images);
            Assert.Equal(ImageFormat.Png, sheet.Images[0].Format);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void WorksheetImage_Constructor_EmptyData_ThrowsException()
    {
        byte[] emptyData = [];
        
        var ex = Assert.Throws<ArgumentException>(() => 
            new WorksheetImage(emptyData, ImageFormat.Png, new(0, 0), 100, 100));
        Assert.Contains("cannot be empty", ex.Message);
    }

    [Fact]
    public void WorksheetImage_Constructor_ZeroWidth_ThrowsException()
    {
        var imageData = CreateTestImageData();
        
        var ex = Assert.Throws<ArgumentException>(() => 
            new WorksheetImage(imageData, ImageFormat.Png, new(0, 0), 0, 100));
        Assert.Contains("Width", ex.Message);
    }

    [Fact]
    public void WorksheetImage_Constructor_ZeroHeight_ThrowsException()
    {
        var imageData = CreateTestImageData();
        
        var ex = Assert.Throws<ArgumentException>(() => 
            new WorksheetImage(imageData, ImageFormat.Png, new(0, 0), 100, 0));
        Assert.Contains("Height", ex.Message);
    }

    [Fact]
    public void WorksheetImage_FromFile_NonexistentFile_ThrowsException()
    {
        var ex = Assert.Throws<FileNotFoundException>(() => 
            WorksheetImage.FromFile("/nonexistent/file.png", new(0, 0), 100, 100));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void WorksheetImage_FromFile_UnsupportedFormat_ThrowsException()
    {
        var tempFile = Path.GetTempFileName() + ".bmp";
        
        try
        {
            File.WriteAllBytes(tempFile, CreateTestImageData());
            
            var ex = Assert.Throws<ArgumentException>(() => 
                WorksheetImage.FromFile(tempFile, new(0, 0), 100, 100));
            Assert.Contains("Unsupported image format", ex.Message);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void WorksheetImage_FromFile_PngFormat_DetectsCorrectly()
    {
        var tempFile = Path.GetTempFileName() + ".png";
        
        try
        {
            File.WriteAllBytes(tempFile, CreateTestImageData());
            
            var image = WorksheetImage.FromFile(tempFile, new(0, 0), 100, 100);
            
            Assert.Equal(ImageFormat.Png, image.Format);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void WorksheetImage_FromFile_JpegFormat_DetectsCorrectly()
    {
        var tempFile = Path.GetTempFileName() + ".jpg";
        
        try
        {
            File.WriteAllBytes(tempFile, CreateTestImageData());
            
            var image = WorksheetImage.FromFile(tempFile, new(0, 0), 100, 100);
            
            Assert.Equal(ImageFormat.Jpeg, image.Format);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void AddImage_MultipleImages_AllAdded()
    {
        var sheet = new WorkSheet("TestSheet");
        var image1 = new WorksheetImage(CreateTestImageData(), ImageFormat.Png, new(0, 0), 100, 100);
        var image2 = new WorksheetImage(CreateTestImageData(), ImageFormat.Jpeg, new(5, 5), 200, 150);
        
        sheet.AddImage(image1);
        sheet.AddImage(image2);
        
        Assert.Equal(2, sheet.Images.Count);
    }

    [Fact]
    public void AddImage_SaveToFile_CreatesValidFile()
    {
        var sheet = new WorkSheet("TestSheet");
        sheet.AddCell(new(0, 0), "Test with Image", null);
        
        var imageData = CreateTestImageData();
        var image = new WorksheetImage(imageData, ImageFormat.Png, new(2, 2), 100, 100);
        sheet.AddImage(image);
        
        var binary = SheetConverter.ToBinaryExcelFile(sheet);
        
        Assert.NotEmpty(binary);
    }

    [Fact]
    public void Images_DefaultValue_IsEmpty()
    {
        var sheet = new WorkSheet("TestSheet");
        
        Assert.Empty(sheet.Images);
    }
}
