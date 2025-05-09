using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Text processing methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Process text replacements in a slide using DollarSignEngine
    /// </summary>
    private void ProcessTextReplacements(SlidePart slidePart)
    {
        // Get all shape elements in the slide
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        Logger.Debug($"Processing text replacements in {shapes.Count} shapes");

        bool hasTextChanges = false;

        foreach (var shape in shapes)
        {
            // Get the shape name
            string shapeName = shape.GetShapeName();
            Logger.Debug($"Processing shape: {shapeName ?? "(unnamed)"}");

            // Update the shape context
            UpdateShapeContext(shape);

            // Get all paragraphs and text runs
            var textRuns = shape.Descendants<A.Text>().ToList();
            Logger.Debug($"Found {textRuns.Count} text runs in shape");

            bool shapeModified = false;

            foreach (var textRun in textRuns)
            {
                string originalText = textRun.Text;
                if (string.IsNullOrEmpty(originalText))
                    continue;

                Logger.Debug($"Processing text: '{originalText}'");

                // Check for variables and PowerPoint functions according to PPT syntax
                if (HasVariablesOrFunctions(originalText))
                {
                    try
                    {
                        // Handle PowerPoint special functions
                        if (originalText.Contains("${ppt."))
                        {
                            shapeModified = ProcessPowerPointFunction(shape, textRun) || shapeModified;
                        }
                        // Handle regular variable replacements using DollarSignEngine
                        else if (originalText.Contains("${"))
                        {
                            string newText = ProcessTextWithVariables(originalText);

                            // Update text if changed
                            if (newText != originalText)
                            {
                                Logger.Debug($"Replacing text: '{originalText}' -> '{newText}'");
                                textRun.Text = newText;
                                shapeModified = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"Error processing text: {originalText}", ex);
                        textRun.Text = $"[Error: {ex.Message}]";
                        shapeModified = true;
                    }
                }
            }

            // Shape was modified, mark for slide save
            if (shapeModified)
            {
                hasTextChanges = true;
            }
        }

        // Save slide after all text replacements if any changes were made
        if (hasTextChanges)
        {
            try
            {
                slidePart.Slide.Save();
                Logger.Debug("Slide saved after text replacements");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error saving slide after text replacements: {ex.Message}", ex);
            }
        }
    }

    /// <summary>
    /// Check if text contains variables or functions according to PPT syntax
    /// </summary>
    private bool HasVariablesOrFunctions(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        // According to PPT syntax, we only check for ${...} pattern
        return text.Contains("${");
    }

    /// <summary>
    /// Process text with variables using DollarSignEngine
    /// </summary>
    private string ProcessTextWithVariables(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        try
        {
            // Prepare variables dictionary
            var variables = PrepareVariablesDictionary();

            // Use DollarSignEngine to evaluate the text with variables
            var result = _expressionEvaluator.Evaluate(text, variables);
            return result?.ToString() ?? text;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing text variables in: {text}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// Process PowerPoint special functions in text using DollarSignEngine
    /// </summary>
    private bool ProcessPowerPointFunction(P.Shape shape, A.Text textRun)
    {
        string text = textRun.Text;
        Logger.Debug($"Processing PowerPoint function: {text}");
        bool textModified = false;

        try
        {
            // Prepare variables dictionary
            var variables = PrepareVariablesDictionary();

            // Extract all ppt. function expressions from the text
            var matches = Regex.Matches(text, @"\${ppt\.(\w+)\(([^)]*)\)}");

            if (matches.Count > 0)
            {
                // If the entire text is a single function call
                if (matches.Count == 1 && matches[0].Value == text)
                {
                    string functionName = matches[0].Groups[1].Value;
                    string parametersString = matches[0].Groups[2].Value;

                    Logger.Debug($"Function: {functionName}, Parameters: {parametersString}");

                    // Execute the function if it exists
                    if (_context.Functions.TryGetValue(functionName, out var function))
                    {
                        // Update shape context
                        _context.Shape.ShapeObject = shape;

                        // Parse parameters
                        var parameters = ParseFunctionParameters(parametersString);

                        // Call the function handler
                        Logger.Debug($"Executing function {functionName} with parameters: {string.Join(", ", parameters)}");
                        var result = function.Execute(_context, null, parameters);

                        // Handle function results
                        if (result is string resultText)
                        {
                            if (string.IsNullOrEmpty(resultText))
                            {
                                // Success case (e.g. image successfully processed)
                                textRun.Text = "";
                                Logger.Debug($"Function {functionName} executed successfully with empty result");
                            }
                            else
                            {
                                // Result text or error message
                                textRun.Text = resultText;
                                Logger.Debug($"Function {functionName} result: {resultText}");
                            }
                            textModified = true;
                        }
                    }
                    else
                    {
                        Logger.Warning($"Function not found: {functionName}");
                        textRun.Text = $"[Unknown function: {functionName}]";
                        textModified = true;
                    }
                }
                // If the text contains multiple expressions or mixed content
                else
                {
                    // Use DollarSignEngine to evaluate the entire text
                    var result = _expressionEvaluator.Evaluate(text, variables);
                    textRun.Text = result?.ToString() ?? "";
                    textModified = true;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing PowerPoint function: {text}", ex);
            textRun.Text = $"[Error: {ex.Message}]";
            textModified = true;
        }

        return textModified;
    }

    /// <summary>
    /// Parse function parameters according to PPT syntax guidelines
    /// </summary>
    private string[] ParseFunctionParameters(string parametersString)
    {
        if (string.IsNullOrEmpty(parametersString))
            return Array.Empty<string>();

        var results = new List<string>();
        bool inQuotes = false;
        int currentStart = 0;
        int parenDepth = 0;

        for (int i = 0; i < parametersString.Length; i++)
        {
            char c = parametersString[i];

            // Handle quotes (start/end of quoted string)
            if (c == '"' && (i == 0 || parametersString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            // Handle nested parentheses
            else if (!inQuotes && c == '(')
            {
                parenDepth++;
            }
            else if (!inQuotes && c == ')')
            {
                parenDepth--;
            }
            // Parameter separator (only at top level, not in quotes or nested parens)
            else if (c == ',' && !inQuotes && parenDepth == 0)
            {
                results.Add(parametersString.Substring(currentStart, i - currentStart).Trim());
                currentStart = i + 1;
            }
        }

        // Add the last parameter
        results.Add(parametersString.Substring(currentStart).Trim());

        // Clean up quoted strings and handle named parameters
        for (int i = 0; i < results.Count; i++)
        {
            var param = results[i].Trim();

            // Handle named parameters (param: value)
            if (param.Contains(":") && !inQuotes)
            {
                var parts = param.Split(new[] { ':' }, 2);
                string paramName = parts[0].Trim();
                string paramValue = parts[1].Trim();

                // If the parameter value is quoted, remove the quotes
                if (paramValue.StartsWith("\"") && paramValue.EndsWith("\"") && paramValue.Length > 1)
                {
                    paramValue = paramValue.Substring(1, paramValue.Length - 2)
                        .Replace("\\\"", "\"")
                        .Replace("\\\\", "\\")
                        .Replace("\\n", "\n")
                        .Replace("\\r", "\r");
                }

                results[i] = $"{paramName}: {paramValue}";
            }
            // Handle regular quoted strings
            else if (param.StartsWith("\"") && param.EndsWith("\"") && param.Length > 1)
            {
                // Remove surrounding quotes and handle escaped characters
                param = param.Substring(1, param.Length - 2)
                    .Replace("\\\"", "\"")
                    .Replace("\\\\", "\\")
                    .Replace("\\n", "\n")
                    .Replace("\\r", "\r");

                results[i] = param;
            }
        }

        Logger.Debug($"Parsed parameters: {string.Join(", ", results)}");
        return results.ToArray();
    }
}