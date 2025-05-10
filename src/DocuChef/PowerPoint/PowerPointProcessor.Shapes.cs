namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Shape related methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Find shapes by name with improved matching logic
    /// </summary>
    private List<P.Shape> FindShapesByName(SlidePart slidePart, string targetName)
    {
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        Logger.Debug($"Looking for shape '{targetName}' among {shapes.Count} shapes");

        var targetShapes = new List<P.Shape>();

        foreach (var shape in shapes)
        {
            // Get shape name using extension method or direct properties
            string shapeName = shape.GetShapeName();
            if (!string.IsNullOrEmpty(shapeName) && shapeName == targetName)
            {
                targetShapes.Add(shape);
                Logger.Debug($"Match found for shape name '{targetName}'");
                continue;
            }

            // Try to check NonVisualDrawingProperties directly
            var nvdp = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
            if (nvdp?.Name?.Value == targetName || nvdp?.Title?.Value == targetName)
            {
                targetShapes.Add(shape);
                Logger.Debug($"Match found for shape via NonVisualDrawingProperties '{targetName}'");
                continue;
            }

            // Try to check alt text through ApplicationNonVisualDrawingProperties attributes
            var anvdp = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;
            if (anvdp != null)
            {
                var descAttr = anvdp.GetAttributes()
                    .FirstOrDefault(a => a.LocalName.Equals("descr", StringComparison.OrdinalIgnoreCase));

                if (descAttr.Value == targetName)
                {
                    targetShapes.Add(shape);
                    Logger.Debug($"Match found for shape via Alt Text '{targetName}'");
                    continue;
                }

                var nameAttr = anvdp.GetAttributes()
                    .FirstOrDefault(a => a.LocalName.Equals("name", StringComparison.OrdinalIgnoreCase));

                if (nameAttr.Value == targetName)
                {
                    targetShapes.Add(shape);
                    Logger.Debug($"Match found for shape via Name attribute '{targetName}'");
                    continue;
                }
            }
        }

        return targetShapes;
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
                var matches = System.Text.RegularExpressions.Regex.Matches(text, @"\${ppt\.(\w+)\(([^)]*)\)}");
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
}