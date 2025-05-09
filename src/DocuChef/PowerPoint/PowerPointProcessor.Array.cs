using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Array handling methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Analyze slide for array indices and automatically handle slide duplication if needed
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

            // We need to duplicate this slide
            DuplicateSlideForArray(presentationPart, slidePart, arrayName, items, itemsPerSlide, slidesNeeded, slideIndex);
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
    /// Duplicate a slide for handling array data across multiple slides
    /// </summary>
    private void DuplicateSlideForArray(PresentationPart presentationPart, SlidePart templateSlidePart,
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

        // Process first slide (original) with the first batch of items
        ProcessSlideWithArrayItems(templateSlidePart, arrayName, items, 0, itemsPerSlide);

        // Create and process additional slides for remaining items
        for (int slideIndex = 1; slideIndex < slidesNeeded; slideIndex++)
        {
            // Calculate start index for this slide's batch
            int batchStartIndex = slideIndex * itemsPerSlide;

            // Clone the template slide
            SlidePart newSlidePart = CloneSlidePart(presentationPart, templateSlidePart);
            string newRelId = presentationPart.GetIdOfPart(newSlidePart);

            // Add new slide to presentation
            P.SlideId newSlideId = new P.SlideId
            {
                Id = maxSlideId + (uint)slideIndex,
                RelationshipId = newRelId
            };

            slideIdList.InsertAt(newSlideId, insertPosition++);

            // Process this slide with its batch of items
            ProcessSlideWithArrayItems(newSlidePart, arrayName, items, batchStartIndex, itemsPerSlide);

            Logger.Debug($"Created and processed slide {slideIndex + 1} for items starting at index {batchStartIndex}");
        }

        // Save the presentation to apply changes
        presentationPart.Presentation.Save();
    }

    /// <summary>
    /// Process a slide with a specific batch of array items
    /// </summary>
    private void ProcessSlideWithArrayItems(SlidePart slidePart, string arrayName, List<object> items, int startIndex, int itemsPerSlide)
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

        // Add the array items with their batch-adjusted indices
        for (int i = 0; i < itemsPerSlide; i++)
        {
            int itemIndex = startIndex + i;
            int localIndex = i;

            // Generate variable key for this item
            string itemKey = $"{arrayName}[{localIndex}]";

            if (itemIndex < items.Count)
            {
                // Item exists in the array
                var item = items[itemIndex];
                batchVariables[itemKey] = item;

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
                                batchVariables[$"{itemKey}.{prop.Name}"] = propValue;
                            }
                            catch
                            {
                                // Skip properties that throw exceptions
                            }
                        }
                    }
                }
            }
            else
            {
                // Item index out of bounds, set null value
                batchVariables[itemKey] = null;
            }
        }

        // Update the context with the batch-specific variables
        var originalVariables = _context.Variables;
        _context.Variables = batchVariables;

        // Process text replacements on this slide
        ProcessTextReplacements(slidePart);

        // Restore original variables
        _context.Variables = originalVariables;
    }

    /// <summary>
    /// Convert an object to a list of objects for array processing
    /// </summary>
    private List<object> ConvertToList(object obj)
    {
        if (obj == null)
            return null;

        if (obj is IList list)
        {
            return list.Cast<object>().ToList();
        }
        else if (obj is IEnumerable enumerable)
        {
            return enumerable.Cast<object>().ToList();
        }

        // Not a collection, return a single-item list
        return new List<object> { obj };
    }
}