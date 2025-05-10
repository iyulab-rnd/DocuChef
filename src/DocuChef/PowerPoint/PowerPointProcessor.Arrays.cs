namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Array handling methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Analyze slide for array indices and handle slide duplication if needed
    /// </summary>
    private void AnalyzeSlideForArrayIndices(PresentationPart presentationPart, SlidePart slidePart, int slideIndex)
    {
        Logger.Debug($"Analyzing slide {slideIndex} for array indices");

        // Find all array index patterns in the slide text
        var arrayReferences = FindArrayReferences(slidePart);
        if (!arrayReferences.Any())
        {
            Logger.Debug("No array references found in slide, no duplication needed");
            return;
        }

        // Group by array name and find max index for each array
        var arrayMaxIndices = arrayReferences
            .GroupBy(r => r.ArrayName)
            .ToDictionary(g => g.Key, g => g.Max(r => r.Index));

        Logger.Debug($"Found {arrayReferences.Count} array references across {arrayMaxIndices.Count} arrays");
        foreach (var entry in arrayMaxIndices)
        {
            Logger.Debug($"Array '{entry.Key}' has max index {entry.Value}");
        }

        // For each array, check if its size exceeds max index and requires duplication
        foreach (var arrayEntry in arrayMaxIndices)
        {
            string arrayName = arrayEntry.Key;
            int maxIndex = arrayEntry.Value;
            int itemsPerSlide = maxIndex + 1; // Calculate items per slide

            // Get array from variables
            object arrayObj = ResolveVariableValue(arrayName);
            if (arrayObj == null)
            {
                Logger.Warning($"Array '{arrayName}' not found in variables");
                continue;
            }

            // Convert to list for processing
            var items = ConvertToList(arrayObj);
            if (items == null || items.Count <= itemsPerSlide)
            {
                Logger.Debug($"Array '{arrayName}' has {items?.Count ?? 0} items, no duplication needed for {itemsPerSlide} items per slide");
                continue;
            }

            // Calculate needed slides
            int slidesNeeded = (int)Math.Ceiling((double)items.Count / itemsPerSlide);
            slidesNeeded = Math.Min(slidesNeeded, _options.MaxSlidesFromTemplate);

            Logger.Info($"Array '{arrayName}' has {items.Count} items, needs {slidesNeeded} slides with {itemsPerSlide} items per slide");

            // Duplicate slides and then process each slide
            DuplicateAndProcessSlidesForArray(presentationPart, slidePart, arrayName, items, itemsPerSlide, slidesNeeded, slideIndex);
        }
    }

    /// <summary>
    /// Find all array references in a slide's text elements
    /// </summary>
    private List<ArrayReference> FindArrayReferences(SlidePart slidePart)
    {
        var result = new List<ArrayReference>();
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();

        // Regular expression to find array indices like ${Items[0].Name} or Items[3]
        var dollarSignRegex = new Regex(@"\${(\w+)\[(\d+)\](\.[\w]+)?}");
        var directRegex = new Regex(@"(\w+)\[(\d+)\](\.[\w]+)?");

        foreach (var shape in shapes)
        {
            var textRuns = shape.Descendants<A.Text>().ToList();
            foreach (var textRun in textRuns)
            {
                if (string.IsNullOrEmpty(textRun.Text))
                    continue;

                // Check for ${array[index].property} pattern
                var dollarMatches = dollarSignRegex.Matches(textRun.Text);
                foreach (Match match in dollarMatches)
                {
                    if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int index))
                    {
                        string arrayName = match.Groups[1].Value;
                        string propPath = match.Groups[3].Success ? match.Groups[3].Value : "";

                        result.Add(new ArrayReference
                        {
                            ArrayName = arrayName,
                            Index = index,
                            PropertyPath = propPath,
                            Pattern = match.Value
                        });
                    }
                }

                // Check for direct array[index].property pattern
                var directMatches = directRegex.Matches(textRun.Text);
                foreach (Match match in directMatches)
                {
                    if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int index))
                    {
                        string arrayName = match.Groups[1].Value;
                        string propPath = match.Groups[3].Success ? match.Groups[3].Value : "";

                        result.Add(new ArrayReference
                        {
                            ArrayName = arrayName,
                            Index = index,
                            PropertyPath = propPath,
                            Pattern = match.Value
                        });
                    }
                }
            }
        }

        return result;
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
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
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
                var newSlideId = new SlideId
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
}