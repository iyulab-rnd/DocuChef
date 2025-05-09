using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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