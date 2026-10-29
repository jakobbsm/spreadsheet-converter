namespace SpreadsheetConverter;

/// <summary>Minimal RFC 4180 writer: a field is quoted only when it contains a comma, quote or line break.</summary>
internal static class Csv
{
    private static readonly System.Buffers.SearchValues<char> CharsNeedingQuotes =
        System.Buffers.SearchValues.Create(",\"\r\n");

    public static void WriteRow(TextWriter writer, string[] fields)
    {
        for (var i = 0; i < fields.Length; i++)
        {
            if (i > 0)
            {
                writer.Write(',');
            }
            WriteField(writer, fields[i]);
        }
        writer.Write('\n');
    }

    private static void WriteField(TextWriter writer, string field)
    {
        if (field.AsSpan().IndexOfAny(CharsNeedingQuotes) < 0)
        {
            writer.Write(field);
            return;
        }

        writer.Write('"');
        writer.Write(field.Replace("\"", "\"\""));
        writer.Write('"');
    }
}
