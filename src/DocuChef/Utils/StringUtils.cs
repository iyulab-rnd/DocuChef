using System.Globalization;
using System.Text.RegularExpressions;

namespace DocuChef.Utils;

/// <summary>
/// Helper methods for string operations
/// </summary>
internal static class StringUtils
{
    /// <summary>
    /// Formats a value using the specified format string and culture
    /// </summary>
    public static string FormatValue(object? value, string? format, CultureInfo culture)
    {
        if (value == null)
            return string.Empty;

        if (string.IsNullOrEmpty(format))
            return value.ToString() ?? string.Empty;

        if (value is IFormattable formattable)
            return formattable.ToString(format, culture);

        return value.ToString() ?? string.Empty;
    }

    /// <summary>
    /// Creates a string representation of a collection
    /// </summary>
    public static string CollectionToString(IEnumerable? collection, string separator = ", ")
    {
        if (collection == null)
            return string.Empty;

        return string.Join(separator, collection.Cast<object>().Select(o => o?.ToString() ?? string.Empty));
    }

    /// <summary>
    /// Removes HTML tags from a string
    /// </summary>
    public static string RemoveHtmlTags(string? html)
    {
        if (string.IsNullOrEmpty(html))
            return string.Empty;

        // Simple HTML tag removal, for more complex scenarios use a proper HTML parser
        return Regex.Replace(html, @"<[^>]*>", string.Empty);
    }

    /// <summary>
    /// Escapes special characters in XML
    /// </summary>
    public static string EscapeXml(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        return text.Replace("&", "&amp;")
                   .Replace("<", "&lt;")
                   .Replace(">", "&gt;")
                   .Replace("\"", "&quot;")
                   .Replace("'", "&apos;");
    }
}