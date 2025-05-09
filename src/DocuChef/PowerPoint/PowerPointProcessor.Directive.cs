using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Directive processing methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Process a slide-level directive (e.g., #slide-foreach)
    /// </summary>
    private void ProcessSlideDirective(PresentationPart presentationPart, SlidePart slidePart, DirectiveContext directive)
    {
        _context.Directive = directive;

        // Handle slide-foreach directive
        if (directive.Name == "slide-foreach")
        {
            ProcessSlideForEach(presentationPart, slidePart, directive);
        }

        // Add support for other slide-level directives as needed
    }

    /// <summary>
    /// Process slide-foreach directive to create multiple slides from template
    /// </summary>
    private void ProcessSlideForEach(PresentationPart presentationPart, SlidePart slidePart, DirectiveContext directive)
    {
        string collectionName = directive.Value;
        string itemName = directive.Parameters.TryGetValue("itemName", out var item) ? item : "item";

        // Resolve the collection from variables
        object collectionObj = ResolveVariableValue(collectionName);
        if (collectionObj == null)
        {
            Logger.Warning($"Collection '{collectionName}' not found for slide-foreach directive");
            return;
        }

        // Handle different collection types
        IEnumerable collection;
        if (collectionObj is IEnumerable enumerable)
        {
            collection = enumerable;
        }
        else
        {
            Logger.Warning($"Value '{collectionName}' is not a collection for slide-foreach directive");
            return;
        }

        Logger.Debug($"Processing slide-foreach with collection '{collectionName}' and item name '{itemName}'");

        // Check if we need to limit the number of slides
        int maxItems = _options.MaxSlidesFromTemplate;
        if (directive.Parameters.TryGetValue("maxItems", out var maxItemsStr) &&
            int.TryParse(maxItemsStr, out int parsedMax))
        {
            maxItems = parsedMax;
        }

        // This implementation is simplified for demonstration. 
        // A complete implementation would:
        // 1. Clone the slide for each item in the collection
        // 2. Update variables for each cloned slide
        // 3. Process text replacements on each cloned slide
        // 4. Handle maxItems, titleTarget, imageTarget parameters

        Logger.Debug($"slide-foreach implementation simplified for demonstration");
    }

    /// <summary>
    /// Process a shape directive (e.g., #foreach, #if)
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
            case "foreach":
                ProcessForEachDirective(slidePart, targetShapes, directive);
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
            // Evaluate the condition using DollarSignEngine
            var result = EvaluateExpression(condition);
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

    /// <summary>
    /// Process foreach directive according to PPT syntax guidelines
    /// </summary>
    private void ProcessForEachDirective(SlidePart slidePart, List<P.Shape> targetShapes, DirectiveContext directive)
    {
        if (targetShapes.Count == 0)
            return;

        string collectionName = directive.Value;
        string itemName = directive.Parameters.TryGetValue("itemName", out var item) ? item : "item";

        // Resolve the collection from variables
        object collectionObj = ResolveVariableValue(collectionName);
        if (collectionObj == null)
        {
            Logger.Warning($"Collection '{collectionName}' not found for foreach directive");
            return;
        }

        // Handle different collection types
        IEnumerable collection;
        if (collectionObj is IEnumerable enumerable)
        {
            collection = enumerable;
        }
        else
        {
            Logger.Warning($"Value '{collectionName}' is not a collection for foreach directive");
            return;
        }

        Logger.Debug($"Processing foreach with collection '{collectionName}' and item name '{itemName}'");

        // Check if we need to limit the number of items
        int maxItems = _options.MaxIterationItems;
        if (directive.Parameters.TryGetValue("maxItems", out var maxItemsStr) &&
            int.TryParse(maxItemsStr, out int parsedMax))
        {
            maxItems = parsedMax;
        }

        // Check if we need special layout handling
        string layout = "vertical";
        if (directive.Parameters.TryGetValue("layout", out var layoutStr))
        {
            layout = layoutStr.ToLowerInvariant();
        }

        // This implementation is simplified.
        // A complete implementation would:
        // 1. Clone the target shape for each item in the collection
        // 2. Position the cloned shapes according to the layout parameter
        // 3. Set context variables for each item and process text replacements
        // 4. Handle maxItems, continueOnNewSlide parameters

        Logger.Debug($"foreach implementation simplified for demonstration");
    }
}