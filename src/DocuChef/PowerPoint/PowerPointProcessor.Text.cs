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

            // Get all shape elements in the slide
            var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
            bool hasChanges = false;

            // Prepare variables for expression evaluation
            var variables = PrepareVariables();

            // Create a FormattedTextProcessor for enhanced handling
            var formattedTextProcessor = new FormattedTextProcessor(this, variables);

            // Process each shape
            foreach (var shape in shapes)
            {
                // Update shape context
                _context.Shape = new ShapeContext
                {
                    Name = shape.GetShapeName(),
                    Id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
                    Text = shape.GetText(),
                    ShapeObject = shape
                };

                Logger.Debug($"Processing shape: {_context.Shape.Name ?? "(unnamed)"}");

                // Process text with formatting preservation
                bool shapeChanged = false;

                // First, try the enhanced FormattedTextProcessor
                if (_options.PreserveTextFormatting)
                {
                    shapeChanged = formattedTextProcessor.ProcessShapeTextWithFormatting(shape);
                    Logger.Debug($"Processed shape with formatting preservation, changed: {shapeChanged}");
                }

                // If no changes or formatting not to be preserved, try SetTextWithExpressions approach
                if (!shapeChanged && _options.PreserveTextFormatting)
                {
                    shapeChanged = shape.SetTextWithExpressions(this, variables);
                    Logger.Debug($"Processed shape with SetTextWithExpressions, changed: {shapeChanged}");
                }

                // If still no changes, try basic method as last resort
                if (!shapeChanged)
                {
                    shapeChanged = ProcessShapeText(shape);
                    Logger.Debug($"Processed shape with basic method, changed: {shapeChanged}");
                }

                if (shapeChanged)
                    hasChanges = true;
            }

            // Save if any changes were made
            if (hasChanges)
            {
                slidePart.Slide.Save();
                Logger.Debug("Slide saved with updated text");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing text replacements: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Process text in a shape with a direct approach (no formatting preservation)
    /// </summary>
    private bool ProcessShapeText(P.Shape shape)
    {
        if (shape.TextBody == null)
            return false;

        try
        {
            // Get complete text from shape
            string completeText = shape.GetText();
            if (string.IsNullOrEmpty(completeText) || !ContainsExpressions(completeText))
                return false;

            // Process the complete text
            string processedText = ProcessExpressions(completeText);
            if (processedText == completeText)
                return false;

            // Update the shape text (this will lose formatting)
            shape.SetText(processedText);
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing shape text: {ex.Message}", ex);
            return false;
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
        return Regex.Replace(text, @"\${([^{}]+)}", match => {
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
    /// Check if text contains expressions
    /// </summary>
    private bool ContainsExpressions(string text)
    {
        return !string.IsNullOrEmpty(text) && text.Contains("${");
    }

    /// <summary>
    /// Parse function parameters
    /// </summary>
    private string[] ParseFunctionParameters(string parametersString)
    {
        if (string.IsNullOrEmpty(parametersString))
            return Array.Empty<string>();

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

        // Clean up parameters
        for (int i = 0; i < results.Count; i++)
        {
            var param = results[i].Trim();

            // Remove quotes from string parameters
            if (param.StartsWith("\"") && param.EndsWith("\"") && param.Length > 1)
            {
                param = param.Substring(1, param.Length - 2);
                results[i] = param;
            }
        }

        return results.ToArray();
    }

    /// <summary>
    /// Replace array references like ${Items[0].Name} in text
    /// </summary>
    private string ReplaceArrayReferences(string text, Dictionary<string, object> variables)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Pattern for ${array[index].property} with optional formatting
        var pattern = @"\$\{(\w+)\[(\d+)\](\.[\w]+)?(:[^}]+)?\}";

        return Regex.Replace(text, pattern, match => {
            try
            {
                string arrayName = match.Groups[1].Value;
                int index = int.Parse(match.Groups[2].Value);
                string propPath = match.Groups[3].Success ? match.Groups[3].Value.Substring(1) : null; // Remove the dot
                string format = match.Groups[4].Success ? match.Groups[4].Value : null;

                // Build the variable key
                string variableKey = propPath != null ?
                    $"{arrayName}[{index}].{propPath}" :
                    $"{arrayName}[{index}]";

                // Look up in variables
                if (variables.TryGetValue(variableKey, out var value))
                {
                    if (value == null)
                        return "";

                    // Apply formatting if specified
                    if (!string.IsNullOrEmpty(format) && format.StartsWith(":") && value is IFormattable formattable)
                    {
                        return formattable.ToString(format.Substring(1), System.Globalization.CultureInfo.CurrentCulture);
                    }

                    return value.ToString();
                }

                return match.Value; // Keep original if not found
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error replacing array reference: {ex.Message}");
                return match.Value;
            }
        });
    }

    /// <summary>
    /// Replace normal variables like ${Variable} in text
    /// </summary>
    private string ReplaceNormalVariables(string text, Dictionary<string, object> variables)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Pattern for ${variable}
        var pattern = @"\$\{([^{}\[\]\.]+)(:[^}]+)?\}";

        return Regex.Replace(text, pattern, match => {
            try
            {
                string variableName = match.Groups[1].Value.Trim();
                string format = match.Groups[2].Success ? match.Groups[2].Value : null;

                // Look up in variables
                if (variables.TryGetValue(variableName, out var value))
                {
                    if (value == null)
                        return "";

                    // Apply formatting if specified
                    if (!string.IsNullOrEmpty(format) && format.StartsWith(":") && value is IFormattable formattable)
                    {
                        return formattable.ToString(format.Substring(1), System.Globalization.CultureInfo.CurrentCulture);
                    }

                    return value.ToString();
                }

                return match.Value; // Keep original if not found
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error replacing variable: {ex.Message}");
                return match.Value;
            }
        });
    }

    /// <summary>
    /// Force text replacement on all shapes in a slide, directly manipulating XML if needed
    /// </summary>
    private void ForceTextReplacementOnSlide(SlidePart slidePart, Dictionary<string, object> variables)
    {
        Logger.Debug("Applying forced text replacement on slide");

        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();

        // Try multiple strategies to ensure text is replaced
        foreach (var shape in shapes)
        {
            string shapeName = shape.GetShapeName();

            if (shape.TextBody == null)
                continue;

            try
            {
                // Strategy 1: Direct XML-based replacement for array items
                var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
                bool directReplacement = false;

                foreach (var paragraph in paragraphs)
                {
                    var textElements = paragraph.Descendants<A.Text>().ToList();

                    // Look for array references in each text element
                    foreach (var textElement in textElements)
                    {
                        if (textElement.Text == null)
                            continue;

                        string originalText = textElement.Text;

                        // Replace array references ${Items[n].Property}
                        string modifiedText = ReplaceArrayReferences(originalText, variables);

                        // Replace normal expressions ${Variable}
                        modifiedText = ReplaceNormalVariables(modifiedText, variables);

                        if (modifiedText != originalText)
                        {
                            textElement.Text = modifiedText;
                            directReplacement = true;
                            Logger.Debug($"Force-updated text in shape {shapeName}: '{originalText}' -> '{modifiedText}'");
                        }
                    }
                }

                // Strategy 2: Combined text replacement if needed
                if (!directReplacement)
                {
                    string completeText = shape.GetText();
                    if (!string.IsNullOrEmpty(completeText))
                    {
                        string replacedText = ReplaceArrayReferences(completeText, variables);
                        replacedText = ReplaceNormalVariables(replacedText, variables);

                        if (replacedText != completeText)
                        {
                            shape.SetText(replacedText);
                            Logger.Debug($"Applied full text replacement in shape {shapeName}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error during forced text replacement in shape {shapeName}: {ex.Message}");
            }
        }

        // Save the slide with changes
        try
        {
            slidePart.Slide.Save();
            Logger.Debug("Saved slide after forced text replacement");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error saving slide after forced text replacement: {ex.Message}");
        }
    }
}