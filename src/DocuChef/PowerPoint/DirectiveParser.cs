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

                // Add special handling for directive-specific parameters
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
        // Specific handling for the 'if' directive
        if (directive.Name == "if")
        {
            // Ensure condition is properly formatted
            directive.Value = directive.Value.Trim();

            // If target is not specified, look for it in parameters
            if (!directive.Parameters.ContainsKey("target"))
            {
                Logger.Warning($"If directive missing required 'target' parameter: {directive.Value}");
            }
        }

        // Add other directive-specific processing as needed
    }
}