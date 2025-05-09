using DocuChef.PowerPoint.Helpers;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
    /// Process PowerPoint functions like ${ppt.Image(...)}
    /// </summary>
    private bool ProcessPowerPointFunctions(P.Shape shape)
    {
        if (shape.TextBody == null)
            return false;

        // Look for PowerPoint functions in all text runs
        bool hasChanges = false;
        var variables = PrepareVariables();

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            foreach (var run in paragraph.Elements<A.Run>().ToList())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;

                // Check for ppt. functions
                if (!text.Contains("${ppt."))
                    continue;

                // Extract function expressions
                var matches = Regex.Matches(text, @"\${ppt\.(\w+)\(([^)]*)\)}");
                if (matches.Count == 0)
                    continue;

                // Process when the entire text is a function call
                if (matches.Count == 1 && matches[0].Value == text)
                {
                    string functionName = matches[0].Groups[1].Value;
                    string parametersString = matches[0].Groups[2].Value;

                    Logger.Debug($"Processing PowerPoint function: {functionName}({parametersString})");

                    // Find the function
                    if (_context.Functions.TryGetValue(functionName, out var function))
                    {
                        // Update context for this shape
                        _context.Shape.ShapeObject = shape;

                        // Parse parameters
                        var parameters = ParseFunctionParameters(parametersString);

                        // Execute function
                        var result = function.Execute(_context, null, parameters);

                        // Handle result
                        if (result is string resultText)
                        {
                            textElement.Text = resultText;
                            hasChanges = true;
                        }
                    }
                    else
                    {
                        Logger.Warning($"Function not found: {functionName}");
                    }
                }
                else
                {
                    // Process mixed content with functions
                    string processedText = ProcessExpressions(text);
                    if (processedText != text)
                    {
                        textElement.Text = processedText;
                        hasChanges = true;
                    }
                }
            }
        }

        return hasChanges;
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
}