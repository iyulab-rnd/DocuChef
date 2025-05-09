using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System.Text.RegularExpressions;

namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Slide cloning methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Simplified clone method to properly maintain OpenXML relationships without circular references
    /// </summary>
    private SlidePart CloneSlideWithRelationships(PresentationPart presentationPart, SlidePart sourceSlidePart)
    {
        // Create a new slide part
        SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();
        Logger.Debug($"Created new slide part: {presentationPart.GetIdOfPart(newSlidePart)}");

        // Copy the slide's XML content
        using (Stream sourceStream = sourceSlidePart.GetStream(FileMode.Open, FileAccess.Read))
        using (Stream targetStream = newSlidePart.GetStream(FileMode.Create, FileAccess.Write))
        {
            sourceStream.CopyTo(targetStream);
        }

        // Track visited parts to prevent circular references
        HashSet<string> visitedPartIds = new HashSet<string>();

        // Clone only specific and important relationships
        CloneSlideRelationships(sourceSlidePart, newSlidePart, visitedPartIds);

        // Save the slide
        newSlidePart.Slide.Save();

        return newSlidePart;
    }

    /// <summary>
    /// Clone only essential slide relationships without creating infinite recursion
    /// </summary>
    private void CloneSlideRelationships(SlidePart sourceSlidePart, SlidePart targetSlidePart, HashSet<string> visitedPartIds)
    {
        // Add this part to visited parts to prevent circular references
        string sourcePartKey = GetPartKey(sourceSlidePart);
        if (visitedPartIds.Contains(sourcePartKey))
        {
            return;
        }
        visitedPartIds.Add(sourcePartKey);

        // Clone essential relationship types
        try
        {
            // 1. SlideLayoutPart - Always needed
            if (sourceSlidePart.SlideLayoutPart != null)
            {
                targetSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
                Logger.Debug($"Reused SlideLayoutPart in cloned slide");
            }

            // 2. NotesSlidePart - Create new but don't recurse
            if (sourceSlidePart.NotesSlidePart != null)
            {
                NotesSlidePart sourceNotesPart = sourceSlidePart.NotesSlidePart;
                NotesSlidePart targetNotesPart = targetSlidePart.AddNewPart<NotesSlidePart>();

                // Only copy content, don't clone relationships of notes
                using (Stream sourceStream = sourceNotesPart.GetStream(FileMode.Open, FileAccess.Read))
                using (Stream targetStream = targetNotesPart.GetStream(FileMode.Create, FileAccess.Write))
                {
                    sourceStream.CopyTo(targetStream);
                }

                Logger.Debug($"Created NotesSlidePart in cloned slide");

                // We don't need to recursively process NotesSlidePart relationships
                // as they're usually not critical for most PowerPoint templates
            }

            // 3. ImageParts - Always clone
            foreach (var idPartPair in sourceSlidePart.Parts)
            {
                if (idPartPair.OpenXmlPart is ImagePart imageSourcePart)
                {
                    ImagePart imageTargetPart = targetSlidePart.AddImagePart(imageSourcePart.ContentType);

                    using (Stream sourceStream = imageSourcePart.GetStream(FileMode.Open, FileAccess.Read))
                    using (Stream targetStream = imageTargetPart.GetStream(FileMode.Create, FileAccess.Write))
                    {
                        sourceStream.CopyTo(targetStream);
                    }

                    // Replace old relationship ID with new one in slide XML if needed
                    // This is only needed in complex scenarios

                    Logger.Debug($"Cloned ImagePart in slide");
                }
            }

            // 4. External relationships (e.g., hyperlinks)
            foreach (var relationship in sourceSlidePart.ExternalRelationships)
            {
                targetSlidePart.AddExternalRelationship(
                    relationship.RelationshipType,
                    relationship.Uri,
                    relationship.Id);

                Logger.Debug($"Copied external relationship: {relationship.Id}");
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error while cloning slide relationships: {ex.Message}");
            // Continue even if some relationships fail
        }
    }

    /// <summary>
    /// Get a unique key for an OpenXmlPart to track visited parts
    /// </summary>
    private string GetPartKey(OpenXmlPart part)
    {
        if (part == null) return string.Empty;

        // Use URI as a unique identifier
        Uri uri = part.Uri;
        if (uri != null)
        {
            return uri.ToString();
        }

        // Fallback to using the hash code
        return part.GetType().Name + part.GetHashCode();
    }

    /// <summary>
    /// Duplicates slides first and then processes each slide with its batch of data
    /// </summary>
    private void DuplicateAndProcessSlidesForArray(
        PresentationPart presentationPart, SlidePart templateSlidePart,
        string arrayName, List<object> items, int itemsPerSlide, int slidesNeeded, int templateIndex)
    {
        // Find template slide in presentation
        var slideIdList = presentationPart.Presentation.SlideIdList;
        var slideIds = slideIdList.ChildElements.OfType<P.SlideId>().ToList();
        int originalIndex = slideIds.FindIndex(id => id.RelationshipId == presentationPart.GetIdOfPart(templateSlidePart));

        if (originalIndex == -1)
        {
            Logger.Error("Template slide not found in presentation");
            return;
        }

        uint maxSlideId = slideIds.Max(id => id.Id.Value);
        int insertPosition = originalIndex + 1;

        Logger.Debug($"Starting slide duplication for array '{arrayName}' with {slidesNeeded} slides needed");

        // First, make copies of the template slide for all needed slides
        List<SlidePart> allSlideParts = new List<SlidePart> { templateSlidePart };

        // Create all the duplicates first (without processing them)
        for (int i = 1; i < slidesNeeded; i++)
        {
            try
            {
                // Clone the original template slide using our enhanced method
                SlidePart newSlidePart = CloneSlideWithRelationships(presentationPart, templateSlidePart);
                string newRelId = presentationPart.GetIdOfPart(newSlidePart);

                // Add new slide to presentation
                var newSlideId = new P.SlideId
                {
                    Id = maxSlideId + (uint)i,
                    RelationshipId = newRelId
                };

                slideIdList.InsertAt(newSlideId, insertPosition++);
                allSlideParts.Add(newSlidePart);

                Logger.Debug($"Created duplicate slide {i + 1} with relationship ID {newRelId}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error creating duplicate slide {i + 1}: {ex.Message}");
                // Continue with other slides even if one fails
            }
        }

        // Save the presentation after adding all slides
        presentationPart.Presentation.Save();

        // Now process each slide with its corresponding data batch
        for (int slideIndex = 0; slideIndex < allSlideParts.Count; slideIndex++)
        {
            try
            {
                // Calculate start index for this slide's batch
                int batchStartIndex = slideIndex * itemsPerSlide;

                // Get the slide part (original or one of the duplicates)
                SlidePart slidePart = allSlideParts[slideIndex];

                // Process this slide with its batch
                Logger.Debug($"Processing slide {slideIndex + 1} with data batch starting at index {batchStartIndex}");
                ProcessSlideWithArrayBatch(slidePart, arrayName, items, batchStartIndex, itemsPerSlide);

                // Save the processed slide
                slidePart.Slide.Save();
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing slide {slideIndex + 1}: {ex.Message}");
                // Continue with other slides even if one fails
            }
        }

        Logger.Debug($"Completed processing all slides for array '{arrayName}'");
    }

    /// <summary>
    /// Improved process for slide with array batch to ensure proper variable replacement
    /// </summary>
    private void ProcessSlideWithArrayBatch(SlidePart slidePart, string arrayName, List<object> items, int startIndex, int itemsPerSlide)
    {
        Logger.Debug($"Processing slide with {items.Count} items, starting at index {startIndex} with {itemsPerSlide} items per slide");

        // Update context for this slide
        _context.SlidePart = slidePart;

        // Create a copy of the variables to avoid modifying the original context
        var batchVariables = new Dictionary<string, object>(_context.Variables);

        // Add batch metadata
        int batchIndex = startIndex / itemsPerSlide;
        batchVariables["_batchIndex"] = batchIndex;
        batchVariables["_batchStartIndex"] = startIndex;
        batchVariables["_batchEndIndex"] = Math.Min(startIndex + itemsPerSlide - 1, items.Count - 1);
        batchVariables["_batchSize"] = Math.Min(itemsPerSlide, items.Count - startIndex);
        batchVariables["_totalItems"] = items.Count;

        // Add the array items with their batch-adjusted indices
        for (int i = 0; i < itemsPerSlide; i++)
        {
            int itemIndex = startIndex + i;
            int localIndex = i;

            // Clear existing variables for this index
            var keysToRemove = batchVariables.Keys
                .Where(k => k.StartsWith($"{arrayName}[{localIndex}]") || k == $"{arrayName}[{localIndex}]")
                .ToList();

            foreach (var key in keysToRemove)
            {
                batchVariables.Remove(key);
            }

            // Generate variable key for this item
            string itemKey = $"{arrayName}[{localIndex}]";

            if (itemIndex < items.Count)
            {
                // Item exists in the array
                var item = items[itemIndex];
                batchVariables[itemKey] = item;

                // Log for debugging
                Logger.Debug($"Setting {itemKey} to item at index {itemIndex} in the source array");

                // If item is an object, also add direct property access
                if (item != null && !item.GetType().IsPrimitive)
                {
                    var properties = item.GetType().GetProperties();
                    foreach (var prop in properties)
                    {
                        if (prop.CanRead)
                        {
                            try
                            {
                                object propValue = prop.GetValue(item);
                                string propKey = $"{itemKey}.{prop.Name}";
                                batchVariables[propKey] = propValue;
                                Logger.Debug($"Setting property {propKey} = {propValue}");
                            }
                            catch (Exception ex)
                            {
                                // Skip properties that throw exceptions
                                Logger.Warning($"Error getting property {prop.Name}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            else
            {
                // Item index out of bounds, set null value
                batchVariables[itemKey] = null;
                Logger.Debug($"Setting {itemKey} to null (index out of bounds)");
            }
        }

        // Update the context with the batch-specific variables
        var originalVariables = _context.Variables;
        _context.Variables = batchVariables;

        try
        {
            // Force replacement directly on shapes - special enhanced processing for batch slides
            ForceTextReplacementOnSlide(slidePart, batchVariables);

            // Standard processing as fallback
            ProcessTextReplacements(slidePart);
            ProcessPowerPointFunctionsInSlide(slidePart);
        }
        finally
        {
            // Restore original variables
            _context.Variables = originalVariables;
        }
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
    /// Helper to copy related parts recursively
    /// </summary>
    private void CopyRelatedParts(OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        foreach (IdPartPair idPartPair in sourcePart.Parts)
        {
            OpenXmlPart childSourcePart = idPartPair.OpenXmlPart;
            string relationshipId = idPartPair.RelationshipId;

            try
            {
                // Try to add the same part
                OpenXmlPart childTargetPart = targetPart.AddPart(childSourcePart, relationshipId);
                Logger.Debug($"Added related part {childSourcePart.GetType().Name} with relId: {relationshipId}");

                // Continue recursively
                CopyRelatedParts(childSourcePart, childTargetPart);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error copying related part {childSourcePart.GetType().Name}: {ex.Message}");
            }
        }

        // Also copy external relationships
        foreach (var relationship in sourcePart.ExternalRelationships)
        {
            targetPart.AddExternalRelationship(
                relationship.RelationshipType,
                relationship.Uri,
                relationship.Id);
        }
    }
}