namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public record ChartSize
{
    public long WidthEMUs { get; init; }
    public long HeightEMUs { get; init; }

    private const long EMUsPerInch = 914400;

    public ChartSize(long widthEMUs, long heightEMUs)
    {
        if (widthEMUs <= 0)
            throw new ArgumentException("Width must be greater than 0", nameof(widthEMUs));
        if (heightEMUs <= 0)
            throw new ArgumentException("Height must be greater than 0", nameof(heightEMUs));

        WidthEMUs = widthEMUs;
        HeightEMUs = heightEMUs;
    }

    public static ChartSize FromEMUs(long width, long height)
    {
        return new(width, height);
    }

    public static ChartSize FromInches(double widthInches, double heightInches)
    {
        if (widthInches <= 0)
            throw new ArgumentException("Width must be greater than 0", nameof(widthInches));
        if (heightInches <= 0)
            throw new ArgumentException("Height must be greater than 0", nameof(heightInches));

        var widthEMUs = (long)(widthInches * EMUsPerInch);
        var heightEMUs = (long)(heightInches * EMUsPerInch);
        return new(widthEMUs, heightEMUs);
    }

    public static ChartSize Default => FromInches(6.5, 4.3);
}
