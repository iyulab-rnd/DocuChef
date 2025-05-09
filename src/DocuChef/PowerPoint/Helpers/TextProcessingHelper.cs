using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Helper class for processing text in PowerPoint slides
/// </summary>
internal class TextProcessingHelper
{
    private readonly PowerPointProcessor _processor;
    private readonly PowerPointContext _context;

    /// <summary>
    /// Initialize text processing helper
    /// </summary>
    public TextProcessingHelper(PowerPointProcessor processor, PowerPointContext context)
    {
        _processor = processor;
        _context = context;
    }

    /// <summary>
    /// Process text replacements in slide - enhanced algorithm with formatting preservation
    /// </summary>
    public void ProcessTextReplacements(SlidePart slidePart, FormattedTextProcessor formattedTextProcessor = null)
    {
        // Get all shape elements in slide
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        Logger.Debug($"Processing text replacements in {shapes.Count} shapes");

        bool hasTextChanges = false;

        foreach (var shape in shapes)
        {
            // Get shape name
            string shapeName = shape.GetShapeName();
            Logger.Debug($"Processing shape: {shapeName ?? "(unnamed)"}");

            // Update shape context
            UpdateShapeContext(shape);

            // Check if shape contains text with expressions
            bool containsExpressions = ContainsTextExpressions(shape);

            if (containsExpressions)
            {
                try
                {
                    bool shapeChanged = false;

                    // First, try to use formattedTextProcessor
                    if (formattedTextProcessor != null)
                    {
                        shapeChanged = formattedTextProcessor.ProcessShapeTextWithFormatting(shape);
                        Logger.Debug($"Processed shape with FormattedTextProcessor, changed: {shapeChanged}");
                    }

                    // If formatting processor didn't work or not provided, fall back to basic processing
                    if (!shapeChanged)
                    {
                        string completeText = ReconstructCompleteText(shape);
                        if (!string.IsNullOrEmpty(completeText) && ContainsExpressions(completeText))
                        {
                            string processedText = ProcessCompleteText(completeText);
                            if (processedText != completeText)
                            {
                                UpdateShapeTextBasic(shape, processedText);
                                shapeChanged = true;
                                Logger.Debug($"Processed shape with basic text processing, text changed");
                            }
                        }
                    }

                    if (shapeChanged)
                        hasTextChanges = true;
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error processing shape text: {ex.Message}", ex);

                    // Last resort fallback - basic text replacement
                    try
                    {
                        string completeText = ReconstructCompleteText(shape);
                        if (!string.IsNullOrEmpty(completeText) && ContainsExpressions(completeText))
                        {
                            string processedText = ProcessCompleteText(completeText);
                            if (processedText != completeText)
                            {
                                ClearShapeText(shape);
                                AddTextToShape(shape, processedText);
                                hasTextChanges = true;
                                Logger.Debug($"Processed shape with emergency fallback processing");
                            }
                        }
                    }
                    catch (Exception fallbackEx)
                    {
                        Logger.Error($"Fallback text processing also failed: {fallbackEx.Message}", fallbackEx);
                    }
                }
            }

            // Check for PowerPoint functions in text
            var powerPointFunctions = FindPowerPointFunctions(shape);
            if (powerPointFunctions.Any())
            {
                Logger.Debug($"Found {powerPointFunctions.Count} PowerPoint functions in shape");

                foreach (var funcMatch in powerPointFunctions)
                {
                    try
                    {
                        bool processed = ProcessPowerPointFunction(shape, funcMatch);
                        if (processed)
                            hasTextChanges = true;
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"Error processing PowerPoint function: {ex.Message}", ex);
                    }
                }
            }
        }

        // Save slide after all text replacements if there were changes
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
    /// Find PowerPoint functions in shape text
    /// </summary>
    private List<Match> FindPowerPointFunctions(P.Shape shape)
    {
        var result = new List<Match>();
        string text = ReconstructCompleteText(shape);

        if (string.IsNullOrEmpty(text))
            return result;

        var pattern = @"\${ppt\.(\w+)\(([^)]*)\)}";
        var matches = Regex.Matches(text, pattern);

        foreach (Match match in matches)
        {
            result.Add(match);
        }

        return result;
    }

    /// <summary>
    /// Process PowerPoint function
    /// </summary>
    private bool ProcessPowerPointFunction(P.Shape shape, Match funcMatch)
    {
        string functionName = funcMatch.Groups[1].Value;
        string parameters = funcMatch.Groups[2].Value;

        Logger.Debug($"Processing PowerPoint function: {functionName} with parameters: {parameters}");

        // Get variables
        var variables = _processor.PrepareVariables();

        // Check if function exists
        if (!variables.TryGetValue($"ppt.{functionName}", out var funcObj) || !(funcObj is PowerPointFunction function))
        {
            Logger.Warning($"Function not found: {functionName}");
            return false;
        }

        // Parse parameters
        var paramList = ParseFunctionParameters(parameters);

        // Update shape context
        _context.Shape.ShapeObject = shape;

        // Execute function
        try
        {
            var result = function.Execute(_context, null, paramList);

            // If result is string, update text
            if (result is string resultText)
            {
                string fullText = ReconstructCompleteText(shape);
                string newText = fullText.Replace(funcMatch.Value, resultText);

                if (newText != fullText)
                {
                    UpdateShapeTextBasic(shape, newText);
                    return true;
                }
            }

            // Function may have modified shape directly (e.g., Image function)
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error executing function {functionName}: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Parse function parameters
    /// </summary>
    private string[] ParseFunctionParameters(string parameters)
    {
        if (string.IsNullOrEmpty(parameters))
            return Array.Empty<string>();

        List<string> result = new List<string>();
        bool inQuotes = false;
        int start = 0;

        for (int i = 0; i < parameters.Length; i++)
        {
            char c = parameters[i];

            if (c == '"' && (i == 0 || parameters[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            else if (c == ',' && !inQuotes)
            {
                result.Add(parameters.Substring(start, i - start).Trim());
                start = i + 1;
            }
        }

        // Add the last parameter
        result.Add(parameters.Substring(start).Trim());

        // Clean up parameters
        for (int i = 0; i < result.Count; i++)
        {
            var param = result[i];

            // Remove quotes from string parameters
            if (param.StartsWith("\"") && param.EndsWith("\""))
            {
                param = param.Substring(1, param.Length - 2);
                result[i] = param;
            }
        }

        return result.ToArray();
    }

    /// <summary>
    /// Check if shape contains text with expressions
    /// </summary>
    private bool ContainsTextExpressions(P.Shape shape)
    {
        if (shape.TextBody == null)
            return false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement != null && ContainsExpressions(textElement.Text))
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Reconstruct complete text from a shape
    /// </summary>
    private string ReconstructCompleteText(P.Shape shape)
    {
        // Get all text runs
        var paragraphs = shape.Descendants<A.Paragraph>().ToList();
        if (paragraphs.Count == 0)
            return string.Empty;

        StringBuilder sb = new StringBuilder();

        foreach (var paragraph in paragraphs)
        {
            StringBuilder paragraphText = new StringBuilder();

            // Merge all text runs in paragraph
            foreach (var run in paragraph.Descendants<A.Run>())
            {
                var text = run.Descendants<A.Text>().FirstOrDefault();
                if (text != null && !string.IsNullOrEmpty(text.Text))
                {
                    paragraphText.Append(text.Text);
                }
            }

            // Add line break between paragraphs
            if (sb.Length > 0 && paragraphText.Length > 0)
            {
                sb.AppendLine();
            }

            sb.Append(paragraphText);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Clear all text from shape
    /// </summary>
    private void ClearShapeText(P.Shape shape)
    {
        if (shape.TextBody == null)
            return;

        var textBody = shape.TextBody;

        // Remove all existing paragraphs
        var paragraphs = textBody.Elements<A.Paragraph>().ToList();
        foreach (var para in paragraphs)
        {
            para.Remove();
        }
    }

    /// <summary>
    /// Add text to shape
    /// </summary>
    private void AddTextToShape(P.Shape shape, string text)
    {
        if (shape.TextBody == null)
            return;

        var textBody = shape.TextBody;

        // Split text by line breaks
        string[] lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0)
            lines = new[] { text };

        // Create a paragraph for each line
        foreach (var line in lines)
        {
            var para = new A.Paragraph();
            var run = new A.Run();
            run.AppendChild(new A.Text(line));
            para.AppendChild(run);
            textBody.AppendChild(para);
        }
    }

    /// <summary>
    /// Update shape text - basic implementation without formatting preservation
    /// </summary>
    private void UpdateShapeTextBasic(P.Shape shape, string newText)
    {
        if (shape.TextBody == null)
            return;

        try
        {
            ClearShapeText(shape);
            AddTextToShape(shape, newText);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error updating shape text: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Process complete text - evaluate all expressions
    /// </summary>
    private string ProcessCompleteText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        try
        {
            // Get prepared variables dictionary
            var variables = _processor.PrepareVariables();

            // Add context variables
            variables["_context"] = _context;

            // 1. Process ${...} format expressions
            var dollarExprPattern = @"\${([^{}]+)}";
            text = Regex.Replace(text, dollarExprPattern, match => {
                string expr = match.Value;
                try
                {
                    var result = _processor.EvaluateCompleteExpression(expr, variables);
                    Logger.Debug($"Evaluated expression '{expr}' to '{result}'");
                    return result?.ToString() ?? "";
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Error evaluating expression '{expr}': {ex.Message}");
                    return expr; // Keep original on error
                }
            });

            return text;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing text: {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// Update shape context
    /// </summary>
    private void UpdateShapeContext(P.Shape shape)
    {
        string name = shape.GetShapeName();
        string text = shape.GetText();

        _context.Shape = new ShapeContext
        {
            Name = name,
            Id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
            Text = text,
            Type = GetShapeType(shape),
            ShapeObject = shape
        };
    }

    /// <summary>
    /// Get shape type
    /// </summary>
    private string GetShapeType(P.Shape shape)
    {
        // Check shape properties
        if (shape.ShapeProperties != null)
        {
            // Look for PresetGeometry
            var presetGeometry = shape.ShapeProperties.ChildElements
                                    .OfType<A.PresetGeometry>()
                                    .FirstOrDefault();

            if (presetGeometry?.Preset != null)
            {
                return presetGeometry.Preset.Value.ToString();
            }
        }

        // Check for TextBody
        if (shape.TextBody != null)
        {
            return "TextBox";
        }

        return "Shape";
    }

    /// <summary>
    /// Check if text contains expressions or functions
    /// </summary>
    private bool ContainsExpressions(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        // Check for ${...} pattern
        return text.Contains("${");
    }
}