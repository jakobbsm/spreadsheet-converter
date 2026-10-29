namespace SpreadsheetConverter;

internal static class ExcelLimits
{
    /// <summary>Rows per worksheet, header included.</summary>
    public const int MaxRows = 1_048_576;

    /// <summary>Columns per worksheet: A to XFD.</summary>
    public const int MaxColumns = 16_384;

    /// <summary>Characters (UTF-16 code units) per cell.</summary>
    public const int MaxCellChars = 32_767;

    /// <summary>Characters in a worksheet name.</summary>
    public const int MaxSheetNameChars = 31;

    /// <summary>Returns at most <paramref name="maxChars"/> characters of <paramref name="value"/>, never splitting a surrogate pair.</summary>
    public static string Clip(string value, int maxChars = MaxCellChars)
    {
        if (value.Length <= maxChars)
        {
            return value;
        }

        var length = maxChars;
        if (char.IsHighSurrogate(value[length - 1]))
        {
            length--;
        }
        return value[..length];
    }

    /// <summary>The column letters for a 1-based column number: 1 is A, 27 is AA.</summary>
    public static string ColumnName(int column)
    {
        var name = "";
        while (column > 0)
        {
            column--;
            name = (char)('A' + column % 26) + name;
            column /= 26;
        }
        return name;
    }
}
