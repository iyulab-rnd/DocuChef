using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint;

/// <summary>
/// Text processing methods for PowerPointProcessor
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Process text replacements in slide with formatting preservation
    /// </summary>
    private void ProcessTextReplacements(SlidePart slidePart)
    {
        try
        {
            Logger.Debug($"Processing text replacements in slide {slidePart.Uri}");

            // Prepare variables for expression evaluation
            var variables = PrepareVariables();

            // Create a FormattedTextProcessor for enhanced handling
            var formattedTextProcessor = new FormattedTextProcessor(this, variables);

            // Use the TextProcessingHelper to process text replacements
            _textHelper.ProcessTextReplacements(slidePart, formattedTextProcessor);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing text replacements: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Process expressions in text
    /// </summary>
    private string ProcessExpressions(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Prepare variables
        var variables = PrepareVariables();

        // Process ${...} expressions
        return Regex.Replace(text, @"\${([^{}]+)}", match =>
        {
            try
            {
                var expressionValue = EvaluateCompleteExpression(match.Value, variables);
                Logger.Debug($"Evaluated expression '{match.Value}' to '{expressionValue}'");
                return expressionValue?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error evaluating expression '{match.Value}': {ex.Message}");
                return match.Value; // Keep original on error
            }
        });
    }

    /// <summary>
    /// Parse function parameters
    /// </summary>
    private string[] ParseFunctionParameters(string parametersString)
    {
        if (string.IsNullOrEmpty(parametersString))
            return Array.Empty<string>();

        // Debug the input
        Logger.Debug($"[PARAM-DEBUG] Parsing parameters: '{parametersString}'");

        // Handle Items[n].Property pattern specially
        var arrayMatch = System.Text.RegularExpressions.Regex.Match(parametersString, @"^Items\[(\d+)\]\.(\w+)$");
        if (arrayMatch.Success)
        {
            int index = int.Parse(arrayMatch.Groups[1].Value);
            string propertyName = arrayMatch.Groups[2].Value;

            Logger.Debug($"[PARAM-DEBUG] Found array pattern: Items[{index}].{propertyName}");

            // Check if index is reasonable
            if (index > 100)
            {
                Logger.Warning($"[PARAM-DEBUG] Suspicious high index: {index}");
            }

            // Just pass as is, actual resolution happens in ImageFunction
            return new string[] { parametersString.Trim() };
        }

        // Regular parsing for other cases
        var results = new List<string>();
        bool inQuotes = false;
        int currentStart = 0;

        for (int i = 0; i < parametersString.Length; i++)
        {
            char c = parametersString[i];

            // Handle quotes
            if (c == '"' && (i == 0 || parametersString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            // Handle parameter separators
            else if (c == ',' && !inQuotes)
            {
                results.Add(parametersString.Substring(currentStart, i - currentStart).Trim());
                currentStart = i + 1;
            }
        }

        // Add the last parameter
        results.Add(parametersString.Substring(currentStart).Trim());

        // Clean up parameters - remove quotes from string literals
        for (int i = 0; i < results.Count; i++)
        {
            var param = results[i].Trim();

            // Remove quotes from string parameters
            if (param.StartsWith("\"") && param.EndsWith("\"") && param.Length > 1)
            {
                param = param.Substring(1, param.Length - 2);
                results[i] = param;
            }

            // Debug each parameter
            Logger.Debug($"[PARAM-DEBUG] Parameter {i}: '{results[i]}'");
        }

        return results.ToArray();
    }
}