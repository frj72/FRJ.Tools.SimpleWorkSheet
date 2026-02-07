using System.Text;
using System.Text.Json;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorkSheet.Components.Import;

public static class GenericTableToJsonConverter
{
    public static string ToJson(GenericTable table)
    {
        ArgumentNullException.ThrowIfNull(table);

        using var stream = new MemoryStream();
        ToJson(table, stream);
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    public static void ToJsonFile(GenericTable table, string filePath)
    {
        ArgumentNullException.ThrowIfNull(table);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        using var stream = File.Create(filePath);
        ToJson(table, stream);
    }

    public static void ToJson(GenericTable table, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(table);
        ArgumentNullException.ThrowIfNull(stream);

        using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true });

        writer.WriteStartArray();

        for (var rowIndex = 0; rowIndex < table.RowCount; rowIndex++)
        {
            writer.WriteStartObject();

            for (var colIndex = 0; colIndex < table.ColumnCount; colIndex++)
            {
                var header = table.GetHeader(colIndex);
                var cellValue = table.GetValue(colIndex, rowIndex);

                writer.WritePropertyName(header);

                if (cellValue == null)
                    writer.WriteNullValue();
                else
                    WriteCellValue(writer, cellValue);
            }

            writer.WriteEndObject();
        }

        writer.WriteEndArray();
        writer.Flush();
    }

    private static void WriteCellValue(Utf8JsonWriter writer, CellValue cellValue) =>
        cellValue.Value.Switch(
            writer.WriteNumberValue,
            writer.WriteNumberValue,
            writer.WriteStringValue,
            dateTimeValue => writer.WriteStringValue(dateTimeValue.ToString("o")),
            dateTimeOffsetValue => writer.WriteStringValue(dateTimeOffsetValue.ToString("o")),
            formulaValue => writer.WriteStringValue(formulaValue.Value)
        );
}
