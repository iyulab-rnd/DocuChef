using System.Text.RegularExpressions;

namespace DocuChef.PowerPoint;

/// <summary>
/// Parses directives from PowerPoint slide notes according to PPT syntax guidelines
/// </summary>
internal static class DirectiveParser
{
    // Regex pattern for matching directives according to PPT syntax guidelines
    // Format: #directive: value, param1: value1, param2: value2
    private static readonly string DefaultPattern = @"#(\w+(?:-\w+)?):([^,]+)(?:,\s*(.+))?";

    /// <summary>
    /// Parse directives from slide notes using the default pattern
    /// </summary>
    public static List<DirectiveContext> ParseDirectives(string notes)
    {
        return ParseDirectives(notes, DefaultPattern);
    }

    /// <summary>
    /// Parse directives from slide notes using a custom pattern
    /// </summary>
    public static List<DirectiveContext> ParseDirectives(string notes, string pattern)
    {
        var directives = new List<DirectiveContext>();

        if (string.IsNullOrEmpty(notes))
            return directives;

        try
        {
            // Match all directive patterns
            var matches = Regex.Matches(notes, pattern, RegexOptions.Compiled);

            foreach (Match match in matches)
            {
                var directive = new DirectiveContext
                {
                    Name = match.Groups[1].Value.Trim(),
                    Value = match.Groups[2].Value.Trim(),
                    Parameters = new Dictionary<string, string>()
                };

                Logger.Debug($"Found directive: {directive.Name} with value: {directive.Value}");

                // Parse parameters
                if (match.Groups.Count > 3 && !string.IsNullOrEmpty(match.Groups[3].Value))
                {
                    var paramString = match.Groups[3].Value.Trim();
                    ParseParameters(paramString, directive.Parameters);
                }

                // Add special handling for directive-specific parameters based on PPT syntax
                ProcessDirectiveSpecificParameters(directive);

                directives.Add(directive);
            }
        }
        catch (Exception ex)
        {
            Logger.Error("Error parsing directives", ex);
        }

        return directives;
    }

    /// <summary>
    /// Parse parameters from a parameter string (param1: value1, param2: value2)
    /// </summary>
    private static void ParseParameters(string paramString, Dictionary<string, string> parameters)
    {
        // Split parameters by commas, but not those inside quotes
        var paramParts = SplitParameterString(paramString);

        foreach (var part in paramParts)
        {
            string paramPair = part.Trim();

            // Split by the first colon to separate name and value
            int colonIndex = paramPair.IndexOf(':');
            if (colonIndex > 0)
            {
                string name = paramPair.Substring(0, colonIndex).Trim();
                string value = paramPair.Substring(colonIndex + 1).Trim();

                // Remove quotes if present
                if (value.StartsWith("\"") && value.EndsWith("\"") && value.Length > 1)
                {
                    value = value.Substring(1, value.Length - 2);
                    // Handle escaped characters
                    value = value.Replace("\\\"", "\"").Replace("\\\\", "\\").Replace("\\n", "\n").Replace("\\r", "\r");
                }

                parameters[name] = value;
                Logger.Debug($"Parsed parameter: {name} = {value}");
            }
            else
            {
                Logger.Warning($"Invalid parameter format: {paramPair}");
            }
        }
    }

    /// <summary>
    /// Splits parameter string by commas, respecting quoted strings
    /// </summary>
    private static List<string> SplitParameterString(string paramString)
    {
        var result = new List<string>();
        bool inQuotes = false;
        int startIndex = 0;

        for (int i = 0; i < paramString.Length; i++)
        {
            char c = paramString[i];

            // Handle quotes
            if (c == '"' && (i == 0 || paramString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            // Handle comma separators (only outside of quotes)
            else if (c == ',' && !inQuotes)
            {
                result.Add(paramString.Substring(startIndex, i - startIndex));
                startIndex = i + 1;
            }
        }

        // Add the last part
        if (startIndex < paramString.Length)
        {
            result.Add(paramString.Substring(startIndex));
        }

        return result;
    }

    /// <summary>
    /// Process directive-specific parameters based on PPT syntax guidelines
    /// </summary>
    private static void ProcessDirectiveSpecificParameters(DirectiveContext directive)
    {
        // Handle foreach directive
        if (directive.Name == "foreach")
        {
            // Check for "as" keyword in the value (collection as item)
            var match = Regex.Match(directive.Value, @"(.+?)\s+as\s+(.+)");
            if (match.Success)
            {
                string collectionName = match.Groups[1].Value.Trim();
                string itemName = match.Groups[2].Value.Trim();

                directive.Value = collectionName;
                directive.Parameters["itemName"] = itemName;
                Logger.Debug($"Parsed 'as' clause in foreach: collection={collectionName}, item={itemName}");
            }
            else
            {
                // Auto-determine item name from collection name as per PPT syntax guidelines
                string itemName = DetermineItemNameFromCollection(directive.Value);
                directive.Parameters["itemName"] = itemName;
                Logger.Debug($"Auto-determined item name: {itemName} from collection: {directive.Value}");
            }
        }
    }

    /// <summary>
    /// Determines item name from collection name according to PPT syntax guidelines
    /// </summary>
    private static string DetermineItemNameFromCollection(string collectionName)
    {
        // Remove any whitespace or quotes
        collectionName = collectionName.Trim();
        if (collectionName.StartsWith("\"") && collectionName.EndsWith("\""))
        {
            collectionName = collectionName.Substring(1, collectionName.Length - 2);
        }

        // Get the base name without any path (e.g., "data.Items" -> "Items")
        string baseName = collectionName;
        if (collectionName.Contains("."))
        {
            baseName = collectionName.Split('.').Last();
        }

        // Rule 1: If ends with 's', remove the 's' for singular form
        if (baseName.EndsWith("s", StringComparison.OrdinalIgnoreCase) && baseName.Length > 1)
        {
            // Handle special cases
            if (baseName.EndsWith("ies", StringComparison.OrdinalIgnoreCase))
            {
                // e.g., "Categories" -> "category"
                return baseName.Substring(0, baseName.Length - 3) + "y";
            }
            if (baseName.EndsWith("es", StringComparison.OrdinalIgnoreCase) &&
                (baseName.EndsWith("xes", StringComparison.OrdinalIgnoreCase) ||
                 baseName.EndsWith("ches", StringComparison.OrdinalIgnoreCase) ||
                 baseName.EndsWith("shes", StringComparison.OrdinalIgnoreCase) ||
                 baseName.EndsWith("sses", StringComparison.OrdinalIgnoreCase)))
            {
                // e.g., "Boxes" -> "box"
                return baseName.Substring(0, baseName.Length - 2);
            }

            // Regular case: "Items" -> "item"
            return baseName.Substring(0, baseName.Length - 1);
        }

        // Rule 2: For non-standard plural forms or non-plural names, use lowercase
        return baseName.ToLowerInvariant();
    }
}