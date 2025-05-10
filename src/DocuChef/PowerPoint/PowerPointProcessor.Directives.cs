namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Directive processing methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Process a slide-level directive (e.g., #if)
    /// </summary>
    private void ProcessSlideDirective(PresentationPart presentationPart, SlidePart slidePart, DirectiveContext directive)
    {
        _context.Directive = directive;

        // Handle if directive for slide-level conditions
        if (directive.Name == "if")
        {
            ProcessSlideIfDirective(presentationPart, slidePart, directive);
        }

        // Other slide-level directives can be added here as needed
    }

    /// <summary>
    /// Process a condition expression in directive
    /// </summary>
    private object EvaluateDirectiveCondition(string expression)
    {
        if (string.IsNullOrEmpty(expression))
            return false;

        try
        {
            // Prepare variables dictionary
            var variables = PrepareVariables();

            // Use ExpressionEvaluator to evaluate the expression
            return _expressionEvaluator.Evaluate(expression, variables);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating directive condition '{expression}': {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Process slide-level if directive
    /// </summary>
    private void ProcessSlideIfDirective(PresentationPart presentationPart, SlidePart slidePart, DirectiveContext directive)
    {
        string condition = directive.Value.Trim();
        Logger.Debug($"Processing slide-level if directive with condition: {condition}");

        try
        {
            // Evaluate the condition
            var result = EvaluateDirectiveCondition(condition);
            bool conditionResult = false;

            // Convert result to boolean
            if (result is bool boolValue)
            {
                conditionResult = boolValue;
            }
            else if (result != null)
            {
                conditionResult = Convert.ToBoolean(result);
            }

            Logger.Debug($"Slide condition evaluated to: {conditionResult}");

            // If condition is false, hide this slide
            if (!conditionResult)
            {
                // Hide slide logic would be implemented here
                // This is a placeholder as slide visibility is complex in OpenXML
                Logger.Debug($"Condition is false, slide should be hidden");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing slide if directive: {condition}", ex);
        }
    }

    /// <summary>
    /// Process a shape directive (e.g., #if)
    /// </summary>
    private void ProcessShapeDirective(SlidePart slidePart, DirectiveContext directive)
    {
        _context.Directive = directive;

        // Get target shape name
        if (!directive.Parameters.TryGetValue("target", out var targetName) || string.IsNullOrEmpty(targetName))
        {
            Logger.Warning($"No target shape specified for directive {directive.Name}");
            return;
        }

        // Remove quotes from target name if present
        targetName = targetName.Trim();
        if (targetName.StartsWith("\"") && targetName.EndsWith("\""))
        {
            targetName = targetName.Substring(1, targetName.Length - 2);
        }

        Logger.Debug($"Processing directive {directive.Name} for target shape '{targetName}'");

        // Find target shapes by name
        var targetShapes = FindShapesByName(slidePart, targetName);

        if (!targetShapes.Any())
        {
            Logger.Warning($"Target shape '{targetName}' not found");
            return;
        }

        Logger.Debug($"Found {targetShapes.Count} shapes matching name '{targetName}'");

        // Handle different directive types
        switch (directive.Name)
        {
            case "if":
                ProcessIfDirective(targetShapes, directive);
                break;
            default:
                Logger.Warning($"Unknown directive: {directive.Name}");
                break;
        }
    }

    /// <summary>
    /// Process if directive according to PPT syntax guidelines
    /// </summary>
    private void ProcessIfDirective(List<P.Shape> targetShapes, DirectiveContext directive)
    {
        string condition = directive.Value.Trim();
        Logger.Debug($"Processing if directive with condition: {condition}");

        try
        {
            // Evaluate the condition using EvaluateDirectiveCondition
            var result = EvaluateDirectiveCondition(condition);
            bool conditionResult = false;

            // Convert result to boolean
            if (result is bool boolValue)
            {
                conditionResult = boolValue;
            }
            else if (result != null)
            {
                conditionResult = Convert.ToBoolean(result);
            }

            Logger.Debug($"Condition evaluated to: {conditionResult}");

            // Get visibleWhenFalse parameter
            string visibleWhenFalseName = null;
            directive.Parameters.TryGetValue("visibleWhenFalse", out visibleWhenFalseName);

            // Handle the target shapes based on condition result
            foreach (var shape in targetShapes)
            {
                shape.SetVisibility(conditionResult);
            }

            // Handle visibleWhenFalse shapes if specified
            if (!string.IsNullOrEmpty(visibleWhenFalseName))
            {
                // Find the visibleWhenFalse shapes
                var visibleWhenFalseShapes = FindShapesByName(_context.SlidePart, visibleWhenFalseName);

                // Set visibility opposite to the condition result
                foreach (var shape in visibleWhenFalseShapes)
                {
                    shape.SetVisibility(!conditionResult);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing if directive: {condition}", ex);
        }
    }
}