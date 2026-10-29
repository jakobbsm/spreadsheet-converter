using System.Text.Json;

namespace SpreadsheetConverter;

internal enum ColumnType
{
    Text,
    Number,
    Bool,
    Date,
    DateTime,
}

/// <param name="Format">Excel number format code for the column's cells, or null for the type's default.</param>
internal sealed record Column(ColumnType Type, string? Format = null);

/// <summary>
/// Column types for <c>to-xlsx</c>, read from a JSON object such as
/// <c>{"employees": "number", "revenue": {"type": "number", "format": "#,##0.00"}}</c>.
/// Columns that are not listed are text.
/// </summary>
internal static class Schema
{
    /// <summary>Excel stores at most 255 characters per number format code.</summary>
    private const int MaxFormatChars = 255;

    public static Dictionary<string, Column> Load(string? path)
    {
        var columns = new Dictionary<string, Column>(StringComparer.Ordinal);
        if (path is null)
        {
            return columns;
        }

        using var document = JsonDocument.Parse(File.ReadAllBytes(path));
        if (document.RootElement.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidDataException("schema: expected a JSON object mapping column names to types");
        }

        foreach (var column in document.RootElement.EnumerateObject())
        {
            columns[column.Name] = column.Value.ValueKind switch
            {
                JsonValueKind.String => new Column(ParseType(column.Name, column.Value.GetString())),
                JsonValueKind.Object => ParseColumn(column.Name, column.Value),
                _ => throw new InvalidDataException(
                    $"schema: column '{column.Name}' must be a type name or an object with \"type\" and \"format\""),
            };
        }
        return columns;
    }

    private static Column ParseColumn(string name, JsonElement value)
    {
        string? type = null;
        string? format = null;
        foreach (var property in value.EnumerateObject())
        {
            switch (property.Name)
            {
                case "type" when property.Value.ValueKind == JsonValueKind.String:
                    type = property.Value.GetString();
                    break;
                case "format" when property.Value.ValueKind == JsonValueKind.String:
                    format = property.Value.GetString();
                    break;
                case "type" or "format":
                    throw new InvalidDataException($"schema: column '{name}': \"{property.Name}\" must be a string");
                default:
                    throw new InvalidDataException(
                        $"schema: column '{name}' has unknown key '{property.Name}' (expected type and format)");
            }
        }

        if (format is not null)
        {
            ValidateFormat(name, format);
        }
        return new Column(type is null ? ColumnType.Text : ParseType(name, type), format);
    }

    private static ColumnType ParseType(string name, string? type) => type switch
    {
        "text" => ColumnType.Text,
        "number" => ColumnType.Number,
        "bool" => ColumnType.Bool,
        "date" => ColumnType.Date,
        "datetime" => ColumnType.DateTime,
        _ => throw new InvalidDataException(
            $"schema: column '{name}' has unknown type '{type}' (expected text, number, bool, date or datetime)"),
    };

    /// <summary>
    /// Format codes are passed to Excel as is, and Excel reports the whole file as corrupt over a malformed one. This
    /// only catches the obvious mistakes: an empty or overlong code, and an unclosed quote or bracket.
    /// </summary>
    private static void ValidateFormat(string name, string format)
    {
        if (format.Length is 0 or > MaxFormatChars)
        {
            throw new InvalidDataException($"schema: column '{name}': format must be 1-{MaxFormatChars} characters");
        }

        var inQuotes = false;
        var inBrackets = false;
        for (var i = 0; i < format.Length; i++)
        {
            switch (format[i])
            {
                case '\\' or '_' or '*' when !inQuotes:
                    i++; // The next character is a literal, padding or fill character.
                    break;
                case '"':
                    inQuotes = !inQuotes;
                    break;
                case '[' when !inQuotes:
                    inBrackets = inBrackets
                        ? throw new InvalidDataException($"schema: column '{name}': format has a nested '['")
                        : true;
                    break;
                case ']' when !inQuotes:
                    inBrackets = inBrackets
                        ? false
                        : throw new InvalidDataException($"schema: column '{name}': format has an unmatched ']'");
                    break;
            }
        }
        if (inQuotes || inBrackets)
        {
            throw new InvalidDataException($"schema: column '{name}': format has an unclosed {(inQuotes ? "quote" : "'['")}");
        }
    }
}
