using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Sheet;

public record WorksheetImage
{
    public byte[] ImageData { get; init; }
    public ImageFormat Format { get; init; }
    public CellPosition Position { get; init; }
    public uint WidthInPixels { get; init; }
    public uint HeightInPixels { get; init; }

    public WorksheetImage(byte[] imageData, ImageFormat format, CellPosition position, uint widthInPixels, uint heightInPixels)
    {
        if (imageData.Length == 0)
            throw new ArgumentException("Image data cannot be empty", nameof(imageData));
        
        if (widthInPixels == 0)
            throw new ArgumentException("Width must be greater than zero", nameof(widthInPixels));
        
        if (heightInPixels == 0)
            throw new ArgumentException("Height must be greater than zero", nameof(heightInPixels));
        
        ImageData = imageData;
        Format = format;
        Position = position;
        WidthInPixels = widthInPixels;
        HeightInPixels = heightInPixels;
    }

    public static WorksheetImage FromFile(string filePath, CellPosition position, uint widthInPixels, uint heightInPixels)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Image file not found", filePath);
        
        var imageData = File.ReadAllBytes(filePath);
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        var format = extension switch
        {
            ".png" => ImageFormat.Png,
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            _ => throw new ArgumentException($"Unsupported image format: {extension}", nameof(filePath))
        };
        
        return new(imageData, format, position, widthInPixels, heightInPixels);
    }
}
