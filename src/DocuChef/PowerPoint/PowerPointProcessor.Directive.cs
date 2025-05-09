using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
            // Evaluate the condition using EvaluateDirectiveCondition instead of EvaluateExpression
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

    /// <summary>
    /// Clone a slide part with all its relationships
    /// </summary>
    private SlidePart CloneSlidePart(PresentationPart presentationPart, SlidePart sourceSlidePart)
    {
        // Create new slide part
        SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();
        Logger.Debug($"Created new slide part: {presentationPart.GetIdOfPart(newSlidePart)}");

        // Copy slide content
        using (Stream sourceStream = sourceSlidePart.GetStream())
        {
            sourceStream.Position = 0;
            newSlidePart.FeedData(sourceStream);
        }
        Logger.Debug("Copied slide content to new slide part");

        // Clone related parts (images, charts, etc.)
        var partIds = new Dictionary<string, string>();

        foreach (IdPartPair part in sourceSlidePart.Parts)
        {
            OpenXmlPart sourcePart = part.OpenXmlPart;
            string relId = part.RelationshipId;

            try
            {
                if (sourcePart is ImagePart imageSourcePart)
                {
                    // Handle image parts
                    ImagePart imageTargetPart = newSlidePart.AddImagePart(imageSourcePart.ContentType, relId);
                    CopyPartContent(imageSourcePart, imageTargetPart);
                    partIds[relId] = relId;
                    Logger.Debug($"Cloned ImagePart with relId: {relId}");
                }
                else if (sourcePart is ChartPart chartSourcePart)
                {
                    // Handle chart parts
                    ChartPart chartTargetPart = newSlidePart.AddNewPart<ChartPart>(relId);
                    CopyPartContent(chartSourcePart, chartTargetPart);
                    CloneChartParts(chartSourcePart, chartTargetPart);
                    partIds[relId] = relId;
                    Logger.Debug($"Cloned ChartPart with relId: {relId}");
                }
                else if (sourcePart is NotesSlidePart notesSlidePart)
                {
                    try
                    {
                        // Handle notes slide parts
                        NotesSlidePart newNotesSlidePart = newSlidePart.AddNewPart<NotesSlidePart>(relId);
                        CopyPartContent(notesSlidePart, newNotesSlidePart);
                        partIds[relId] = relId;
                        Logger.Debug($"Cloned NotesSlidePart with relId: {relId}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to clone NotesSlidePart: {ex.Message}");
                    }
                }
                else if (sourcePart is SlideLayoutPart slideLayoutPart)
                {
                    try
                    {
                        // Handle layout parts
                        SlideLayoutPart newLayoutPart = newSlidePart.AddNewPart<SlideLayoutPart>(relId);
                        CopyPartContent(slideLayoutPart, newLayoutPart);
                        partIds[relId] = relId;
                        Logger.Debug($"Cloned SlideLayoutPart with relId: {relId}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to clone SlideLayoutPart: {ex.Message}");
                    }
                }
                else
                {
                    // For other part types, try to use a generic approach
                    Type partType = sourcePart.GetType();
                    Logger.Debug($"Attempting to clone part of type: {partType.Name}");

                    try
                    {
                        // Use reflection to find the appropriate AddNewPart method
                        var addMethod = typeof(SlidePart).GetMethods()
                            .Where(m => m.Name == "AddNewPart" && m.IsGenericMethod)
                            .FirstOrDefault();

                        if (addMethod != null)
                        {
                            var genericMethod = addMethod.MakeGenericMethod(partType);
                            var newPart = genericMethod.Invoke(newSlidePart, new object[] { relId }) as OpenXmlPart;

                            if (newPart != null)
                            {
                                CopyPartContent(sourcePart, newPart);
                                partIds[relId] = relId;
                                Logger.Debug($"Cloned generic part of type {partType.Name} with relId: {relId}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to clone part type {sourcePart.GetType().Name}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error cloning part: {ex.Message}");
            }
        }

        return newSlidePart;
    }

    /// <summary>
    /// Clone chart-related parts
    /// </summary>
    private void CloneChartParts(ChartPart sourceChartPart, ChartPart targetChartPart)
    {
        // Handle embedding package if present
        if (sourceChartPart.EmbeddedPackagePart != null)
        {
            try
            {
                var targetPackagePart = targetChartPart.AddEmbeddedPackagePart("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                CopyPartContent(sourceChartPart.EmbeddedPackagePart, targetPackagePart);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to clone embedded package part: {ex.Message}");
            }
        }

        // Handle other chart-related parts as needed...
    }

    /// <summary>
    /// Copy content from one OpenXmlPart to another with error handling
    /// </summary>
    private void CopyPartContent(OpenXmlPart source, OpenXmlPart target)
    {
        try
        {
            using (Stream sourceStream = source.GetStream(FileMode.Open, FileAccess.Read))
            {
                sourceStream.Position = 0;
                using (Stream targetStream = target.GetStream(FileMode.Create, FileAccess.Write))
                {
                    byte[] buffer = new byte[8192];
                    int bytesRead;
                    while ((bytesRead = sourceStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        targetStream.Write(buffer, 0, bytesRead);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error copying part content: {ex.Message}");

            // Fallback method using FeedData
            try
            {
                using (Stream sourceStream = source.GetStream())
                {
                    sourceStream.Position = 0;
                    target.FeedData(sourceStream);
                }
            }
            catch (Exception feedEx)
            {
                Logger.Warning($"Error using FeedData: {feedEx.Message}");
            }
        }
    }
}