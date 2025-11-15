using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public static class CellImportExtensions
{
    extension(CellBuilder builder)
    {
        public CellBuilder FromImportedValue(string rawValue,
            string source)
        {
            return builder
                .WithMetadata(meta => meta
                    .WithSource(source)
                    .WithOriginalValue(rawValue)
                    .WithImportedAt(DateTime.UtcNow));
        }

        public CellBuilder FromImportedValue(string rawValue,
            ImportOptions options)
        {
            var metadataBuilder = builder;

            if (options.PreserveOriginalValue)
                metadataBuilder = metadataBuilder.WithMetadata(meta => meta
                    .WithOriginalValue(rawValue)
                    .WithImportedAt(DateTime.UtcNow));

            if (!string.IsNullOrEmpty(options.SourceIdentifier))
                metadataBuilder = metadataBuilder.WithMetadata(meta => meta
                    .WithSource(options.SourceIdentifier));

            if (options.CustomMetadata != null) metadataBuilder = options.CustomMetadata.Aggregate(metadataBuilder, (current, kvp) => current.WithMetadata(meta => meta.AddCustomData(kvp.Key, kvp.Value)));

            if (options.DefaultStyle != null) metadataBuilder = metadataBuilder.WithStyle(options.DefaultStyle);

            return metadataBuilder;
        }
    }

    public static string ProcessRawValue(this string rawValue, ImportOptions? options)
    {
        if (options == null)
            return rawValue;

        return options.TrimWhitespace ? rawValue.Trim() : rawValue;
    }
}
