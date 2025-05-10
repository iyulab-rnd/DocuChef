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

        // Create a FormattedTextProcessor if one wasn't provided
        if (formattedTextProcessor == null)
        {
            var variables = _processor.PrepareVariables();
            formattedTextProcessor = new FormattedTextProcessor(_processor, variables);
        }

        foreach (var shape in shapes)
        {
            // Get shape name
            string shapeName = shape.GetShapeName();
            Logger.Debug($"Processing shape: {shapeName ?? "(unnamed)"}");

            // Update shape context
            UpdateShapeContext(shape);

            // Process text with multiple strategies for maximum compatibility
            bool shapeChanged = ProcessShapeWithMultipleStrategies(shape, formattedTextProcessor);

            if (shapeChanged)
                hasTextChanges = true;

            // Process any PowerPoint functions in the shape text
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
    /// Update shape context with current shape information
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
    /// Process a shape with multiple text processing strategies
    /// </summary>
    private bool ProcessShapeWithMultipleStrategies(P.Shape shape, FormattedTextProcessor formattedTextProcessor)
    {
        // Skip shapes without text content
        if (shape.TextBody == null)
            return false;

        // Check if shape contains text with expressions
        bool containsExpressions = ContainsTextExpressions(shape);
        if (!containsExpressions)
            return false;

        try
        {
            // Strategy 1: Try to process with the FormattedTextProcessor (best formatting preservation)
            bool changed = formattedTextProcessor.ProcessShapeTextWithFormatting(shape);
            if (changed)
            {
                Logger.Debug("Successfully processed shape with FormattedTextProcessor");
                return true;
            }

            // Strategy 2: Try using SetTextWithExpressions approach (good formatting preservation)
            var variables = _processor.PrepareVariables();
            changed = shape.SetTextWithExpressions(_processor, variables);
            if (changed)
            {
                Logger.Debug("Successfully processed shape with SetTextWithExpressions");
                return true;
            }

            // Strategy 3: Reconstruct and process the complete text (fallback, might lose formatting)
            string completeText = ReconstructCompleteText(shape);
            if (!string.IsNullOrEmpty(completeText) && ContainsExpressions(completeText))
            {
                string processedText = ProcessCompleteText(completeText);
                if (processedText != completeText)
                {
                    // Try to update while preserving formatting as much as possible
                    try
                    {
                        UpdateShapeTextWithFormatting(shape, processedText);
                    }
                    catch (Exception)
                    {
                        // If that fails, fall back to basic replacement
                        UpdateShapeTextBasic(shape, processedText);
                    }

                    Logger.Debug("Successfully processed shape with text reconstruction method");
                    return true;
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing shape text: {ex.Message}", ex);

            // Emergency fallback - try one last approach
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
                        Logger.Debug("Processed shape with emergency fallback method");
                        return true;
                    }
                }
            }
            catch (Exception fallbackEx)
            {
                Logger.Error($"Fallback text processing also failed: {fallbackEx.Message}", fallbackEx);
            }

            return false;
        }
    }

    /// <summary>
    /// Update shape text while attempting to preserve formatting
    /// </summary>
    private void UpdateShapeTextWithFormatting(P.Shape shape, string newText)
    {
        // Get existing text with formatting information
        var formattedText = shape.GetFormattedText();
        if (formattedText.Count == 0)
        {
            // No formatted text found, fall back to basic update
            UpdateShapeTextBasic(shape, newText);
            return;
        }

        // Clear existing text
        ClearShapeText(shape);

        // Split new text into lines
        string[] lines = newText.Split(new[] { '\r', '\n' }, StringSplitOptions.None);

        var textBody = shape.TextBody;

        // Try to preserve paragraph and run properties from original text
        int paraIndex = 0;
        foreach (var line in lines)
        {
            // Create a new paragraph
            var para = new A.Paragraph();

            // Try to get formatting from original paragraph
            A.RunProperties runProps = null;
            if (paraIndex < formattedText.Count)
            {
                runProps = formattedText[paraIndex].Properties;
            }

            // Create a run with the line text
            var run = new A.Run();
            if (runProps != null)
            {
                run.RunProperties = (A.RunProperties)runProps.CloneNode(true);
            }

            run.AppendChild(new A.Text(line));
            para.AppendChild(run);
            textBody.AppendChild(para);

            paraIndex++;
        }
    }

    /// <summary>
    /// Find PowerPoint functions in shape text
    /// </summary>
    public List<Match> FindPowerPointFunctions(P.Shape shape)
    {
        var result = new List<Match>();
        string text = ReconstructCompleteText(shape);

        if (string.IsNullOrEmpty(text))
            return result;

        // 일반 함수 패턴 (${ppt.Function(...)})
        var pattern = @"\${ppt\.(\w+)\(([^)]*)\)}";
        var matches = Regex.Matches(text, pattern);

        foreach (Match match in matches)
        {
            result.Add(match);
        }

        // 배열 참조가 포함된 함수 패턴 (${ppt.Function(Items[0].Property)})
        var arrayPattern = @"\${ppt\.(\w+)\((\w+)\[(\d+)\][^)]*\)}";
        var arrayMatches = Regex.Matches(text, arrayPattern);

        foreach (Match match in arrayMatches)
        {
            if (!result.Contains(match))
            {
                result.Add(match);
            }
        }

        return result;
    }

    /// <summary>
    /// Process a PowerPoint function
    /// </summary>
    public bool ProcessPowerPointFunction(P.Shape shape, Match funcMatch)
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
    /// Parse function parameters with support for array references
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
        if (start < parameters.Length)
        {
            result.Add(parameters.Substring(start).Trim());
        }

        // Clean up parameters
        for (int i = 0; i < result.Count; i++)
        {
            var param = result[i].Trim();

            // Remove quotes from string parameters
            if (param.StartsWith("\"") && param.EndsWith("\"") && param.Length > 1)
            {
                param = param.Substring(1, param.Length - 2)
                    .Replace("\\\"", "\"")
                    .Replace("\\\\", "\\");
                result[i] = param;
            }
        }

        return result.ToArray();
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

            // Process ${...} format expressions
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
    /// Clear all text from shape
    /// </summary>
    private void ClearShapeText(P.Shape shape)
    {
        if (shape.TextBody == null)
            return;

        var textBody = shape.TextBody;

        // Store paragraph properties for reuse
        var paragraphProperties = textBody.Elements<A.Paragraph>()
            .Select(p => p.ParagraphProperties?.CloneNode(true) as A.ParagraphProperties)
            .ToList();

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