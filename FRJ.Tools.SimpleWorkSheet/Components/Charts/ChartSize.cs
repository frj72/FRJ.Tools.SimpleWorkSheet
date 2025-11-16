namespace FRJ.Tools.SimpleWorkSheet.Components.Charts;

public record ChartSize
{
    public long WidthEmus { get; init; }
    public long HeightEmus { get; init; }

    private const long EmusPerInch = 914400;

    public ChartSize(long widthEmus, long heightEmus)
    {
        if (widthEmus <= 0)
            throw new ArgumentException("Width must be greater than 0", nameof(widthEmus));
        if (heightEmus <= 0)
            throw new ArgumentException("Height must be greater than 0", nameof(heightEmus));

        WidthEmus = widthEmus;
        HeightEmus = heightEmus;
    }

    public static ChartSize FromEmus(long width, long height)
    {
        return new(width, height);
    }

    public static ChartSize FromInches(double widthInches, double heightInches)
    {
        if (widthInches <= 0)
            throw new ArgumentException("Width must be greater than 0", nameof(widthInches));
        if (heightInches <= 0)
            throw new ArgumentException("Height must be greater than 0", nameof(heightInches));

        var widthEmus = (long)(widthInches * EmusPerInch);
        var heightEmus = (long)(heightInches * EmusPerInch);
        return new(widthEmus, heightEmus);
    }

    public static ChartSize Default => FromInches(6.5, 4.3);
}
