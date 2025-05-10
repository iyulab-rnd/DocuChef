namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Shape related methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Find shapes by name
    /// </summary>
    private List<P.Shape> FindShapesByName(SlidePart slidePart, string targetName)
    {
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        var targetShapes = new List<P.Shape>();

        foreach (var shape in shapes)
        {
            string shapeName = shape.GetShapeName();

            if (shapeName == targetName)
            {
                targetShapes.Add(shape);
                Logger.Debug($"Found shape '{targetName}'");
            }
        }

        return targetShapes;
    }

    /// <summary>
    /// Update shape context
    /// </summary>
    private void UpdateShapeContext(P.Shape shape)
    {
        _context.Shape = new ShapeContext
        {
            Name = shape.GetShapeName(),
            Id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
            Text = shape.GetText(),
            Type = GetShapeType(shape),
            ShapeObject = shape
        };
    }

    /// <summary>
    /// Get shape type
    /// </summary>
    private string GetShapeType(P.Shape shape)
    {
        var presetGeometry = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        if (presetGeometry?.Preset != null)
        {
            return presetGeometry.Preset.Value.ToString();
        }

        return shape.TextBody != null ? "TextBox" : "Shape";
    }

    /// <summary>
    /// Process PowerPoint functions in shape
    /// </summary>
    private bool ProcessPowerPointFunctions(P.Shape shape)
    {
        if (shape.TextBody == null)
            return false;

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
                var functions = ExtractPowerPointFunctions(text);

                if (!functions.Any())
                    continue;

                // Process function if it's the entire text
                if (functions.Count == 1 && ExpressionProcessor.IsSingleExpression(text))
                {
                    var function = functions[0];
                    if (ProcessPowerPointFunction(function, shape))
                    {
                        hasChanges = true;
                    }
                }
                else
                {
                    // Process mixed content
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
    /// Extract PowerPoint functions from text
    /// </summary>
    private List<PowerPointFunctionCall> ExtractPowerPointFunctions(string text)
    {
        var result = new List<PowerPointFunctionCall>();
        var pattern = new Regex(@"\${ppt\.(\w+)\(([^)]*)\)}", RegexOptions.Compiled);

        var matches = pattern.Matches(text);
        foreach (Match match in matches)
        {
            result.Add(new PowerPointFunctionCall
            {
                FullMatch = match.Value,
                FunctionName = match.Groups[1].Value,
                Parameters = match.Groups[2].Value
            });
        }

        return result;
    }

    /// <summary>
    /// Process a single PowerPoint function
    /// </summary>
    private bool ProcessPowerPointFunction(PowerPointFunctionCall functionCall, P.Shape shape)
    {
        Logger.Debug($"Processing PowerPoint function: {functionCall.FunctionName}({functionCall.Parameters})");

        // Find the function
        if (!_context.Functions.TryGetValue(functionCall.FunctionName, out var function))
        {
            Logger.Warning($"Function not found: {functionCall.FunctionName}");
            return false;
        }

        // Update shape context
        _context.Shape.ShapeObject = shape;

        // Parse parameters
        var parameters = ParseFunctionParameters(functionCall.Parameters);

        try
        {
            // Execute function
            var result = function.Execute(_context, null, parameters);

            // Handle result
            if (result is string resultText)
            {
                var textElements = shape.Descendants<A.Text>()
                    .Where(t => t.Text == functionCall.FullMatch)
                    .ToList();

                foreach (var textElement in textElements)
                {
                    textElement.Text = resultText;
                }

                return true;
            }

            // Function may have modified shape directly
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error executing function {functionCall.FunctionName}: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Represents a PowerPoint function call
    /// </summary>
    private class PowerPointFunctionCall
    {
        public string FullMatch { get; set; }
        public string FunctionName { get; set; }
        public string Parameters { get; set; }
    }
}