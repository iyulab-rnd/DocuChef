using DocuChef.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocuChef.PowerPoint;

public partial class PowerPointRecipe
{
    private async Task ProcessSlidesAsync()
    {
        // Get presentation part
        var presentationPart = Document?.PresentationPart;
        if (presentationPart == null)
        {
            LoggingHelper.LogWarning("Presentation part is null");
            return;
        }

        // Get all slides
        var presentation = presentationPart.Presentation;
        if (presentation == null)
        {
            LoggingHelper.LogWarning("Presentation is null");
            return;
        }

        var slideIdList = presentation.SlideIdList;
        if (slideIdList == null)
        {
            LoggingHelper.LogWarning("Slide ID list is null");
            return;
        }

        // Improve type compatibility
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();

        var slidesToProcess = new List<(SlideId id, SlidePart part, SlideDirective? directive)>();

        // Analyze slide notes to determine special processing
        foreach (SlideId slideId in slideIds)
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrEmpty(relationshipId))
            {
                continue;
            }

            try
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
                if (slidePart == null) continue;

                var directive = GetSlideDirective(slidePart);
                slidesToProcess.Add((slideId, slidePart, directive));
            }
            catch (Exception ex)
            {
                LoggingHelper.LogError($"Error accessing slide part for slide ID {slideId.Id}", ex);
                throw new TemplateException($"Error accessing slide part: {ex.Message}", ex);
            }
        }

        // Process collection slides (repeat directive)
        var newSlideIds = new List<SlideId>();
        var slidesToDelete = new List<SlideId>();

        foreach (var (slideId, slidePart, directive) in slidesToProcess)
        {
            if (directive == null)
            {
                // Regular slide - just process variables
                await ProcessSlideContentAsync(slidePart);
                newSlideIds.Add(slideId);
                continue;
            }

            if (directive.Type == "repeat")
            {
                // Mark original for deletion
                slidesToDelete.Add(slideId);

                // Get collection to repeat
                var collection = TextProcessingHelper.ResolveCollection(directive.Value, Data, Options.VariableResolver);
                if (!collection.Any()) continue;

                // Create a slide for each collection item
                foreach (var item in collection)
                {
                    try
                    {
                        // Clone slide
                        var newSlide = await CloneSlideAsync(slidePart, presentationPart);
                        await ProcessSlideContentWithContextAsync(newSlide, item);

                        // Add to presentation
                        // Generate unique ID to prevent conflicts
                        uint newId = GenerateUniqueSlideId(presentation);
                        var newSlideId = new SlideId { Id = newId, RelationshipId = presentationPart.GetIdOfPart(newSlide) };
                        newSlideIds.Add(newSlideId);
                    }
                    catch (Exception ex)
                    {
                        LoggingHelper.LogError($"Error processing collection slide for directive {directive.Type}:{directive.Value}", ex);
                        throw new TemplateException($"Error processing collection slide: {ex.Message}", ex);
                    }
                }
            }
            else if (directive.Type == "if")
            {
                // Conditional slide
                bool condition = TextProcessingHelper.EvaluateCondition(directive.Value, Data, Options.VariableResolver);

                if (condition)
                {
                    await ProcessSlideContentAsync(slidePart);
                    newSlideIds.Add(slideId);
                }
                else
                {
                    slidesToDelete.Add(slideId);
                }
            }
        }

        // Update slide order
        var slideIdList2 = presentationPart.Presentation.SlideIdList;
        if (slideIdList2 != null)
        {
            try
            {
                // Remove deleted slides
                foreach (var slideId in slidesToDelete)
                {
                    slideIdList2.RemoveChild(slideId);
                }

                // Clear and rebuild slide list
                slideIdList2.RemoveAllChildren();
                foreach (var slideId in newSlideIds)
                {
                    slideIdList2.AppendChild(slideId);
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.LogError("Error updating slide order", ex);
                throw new TemplateException($"Error updating slide order: {ex.Message}", ex);
            }
        }
    }

    // Generate unique ID to prevent conflicts
    private static uint GenerateUniqueSlideId(Presentation presentation)
    {
        var existingIds = presentation.SlideIdList?.ChildElements
            .OfType<SlideId>()
            .Select(s => s.Id?.Value ?? 0)
            .ToList() ?? [];

        // Start ID value
        uint startId = 256; // PowerPoint slides typically start at 256

        // Find an unused ID
        while (existingIds.Contains(startId))
        {
            startId++;
        }

        return startId;
    }

    // Clone slide asynchronously
    private async Task<SlidePart> CloneSlideAsync(SlidePart sourcePart, PresentationPart presentationPart)
    {
        return await Task.Run(() => {
            // Create new slide part
            var targetSlidePart = presentationPart.AddNewPart<SlidePart>();

            // Clone slide XML
            using (var sourceStream = sourcePart.GetStream())
            using (var targetStream = targetSlidePart.GetStream(FileMode.Create))
            {
                sourceStream.CopyTo(targetStream);
            }

            // Handle relationships
            CopyRelationships(sourcePart, targetSlidePart);

            return targetSlidePart;
        });
    }

    // Copy relationships
    private static void CopyRelationships(SlidePart sourcePart, SlidePart targetPart)
    {
        // Copy image parts
        foreach (var imagePart in sourcePart.ImageParts)
        {
            var imagePartId = sourcePart.GetIdOfPart(imagePart);
            var targetImagePart = targetPart.AddImagePart(imagePart.ContentType);

            using (var stream = imagePart.GetStream())
            using (var targetStream = targetImagePart.GetStream(FileMode.Create))
            {
                stream.CopyTo(targetStream);
            }

            // Handle relationship IDs
            var targetImagePartId = targetPart.GetIdOfPart(targetImagePart);
            UpdateRelationshipIds(targetPart, imagePartId, targetImagePartId);
        }

        // Copy chart parts
        foreach (var chartPart in sourcePart.ChartParts)
        {
            var chartPartId = sourcePart.GetIdOfPart(chartPart);
            var targetChartPart = targetPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();

            using (var stream = chartPart.GetStream())
            using (var targetStream = targetChartPart.GetStream(FileMode.Create))
            {
                stream.CopyTo(targetStream);
            }

            // Copy chart data
            if (chartPart.EmbeddedPackagePart != null)
            {
                var embeddedPackagePart = targetChartPart.AddEmbeddedPackagePart(chartPart.EmbeddedPackagePart.ContentType);
                using var stream = chartPart.EmbeddedPackagePart.GetStream();
                using var targetStream = embeddedPackagePart.GetStream(FileMode.Create);
                stream.CopyTo(targetStream);
            }

            // Handle relationship IDs
            var targetChartPartId = targetPart.GetIdOfPart(targetChartPart);
            UpdateRelationshipIds(targetPart, chartPartId, targetChartPartId);
        }

        // Copy external relationships
        foreach (var relationship in sourcePart.ExternalRelationships)
        {
            targetPart.AddExternalRelationship(
                relationship.RelationshipType,
                relationship.Uri,
                relationship.Id);
        }
    }

    // Update relationship IDs in XML content
    private static void UpdateRelationshipIds(SlidePart targetPart, string oldId, string newId)
    {
        try
        {
            // Update relationship ID references in slide XML
            using var stream = targetPart.GetStream(FileMode.Open);
            var xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.Load(stream);

            // PowerPoint relationship namespace
            var namespaceManager = new System.Xml.XmlNamespaceManager(xmlDoc.NameTable);
            namespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            // Update relationship ID references
            var nodes = xmlDoc.SelectNodes("//*[@r:id]", namespaceManager);
            if (nodes != null)
            {
                foreach (System.Xml.XmlNode node in nodes)
                {
                    var idAttr = node.Attributes?["id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"];
                    if (idAttr != null && idAttr.Value == oldId)
                    {
                        idAttr.Value = newId;
                    }
                }
            }

            // Save changes
            stream.SetLength(0);
            xmlDoc.Save(stream);
        }
        catch (Exception ex)
        {
            // Log error but continue with cloning
            LoggingHelper.LogWarning($"Error updating relationship IDs: {ex.Message}", ex);
        }
    }

    // Update slide numbers
    private async Task UpdateSlideNumbersAsync()
    {
        await Task.Run(() => {
            if (Document?.PresentationPart == null) return;

            var presentationPart = Document.PresentationPart;
            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList == null) return;

            var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();

            // Process each slide
            for (int i = 0; i < slideIds.Count; i++)
            {
                var slideId = slideIds[i];
                var relationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relationshipId)) continue;

                try
                {
                    var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
                    if (slidePart == null) continue;

                    // Find slide number placeholders
                    foreach (var shape in slidePart.Slide.Descendants<Shape>())
                    {
                        var placeholderType = shape.Descendants<PlaceholderShape>()
                            .FirstOrDefault()?.Type?.Value;

                        // If this is a slide number placeholder
                        if (placeholderType == PlaceholderValues.SlideNumber)
                        {
                            // Update text content
                            foreach (var textElement in shape.Descendants<A.Text>())
                            {
                                textElement.Text = (i + 1).ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log error but continue
                    LoggingHelper.LogWarning($"Error updating slide number for slide {i + 1}", ex);
                }
            }
        });
    }

    private SlideDirective? GetSlideDirective(SlidePart slidePart)
    {
        // Check if the slide has notes
        if (slidePart.NotesSlidePart == null) return null;

        // Get notes text
        var notesText = GetNotesText(slidePart.NotesSlidePart);
        if (string.IsNullOrEmpty(notesText)) return null;

        // Look for directives
        var match = SlideNoteRegex.Match(notesText);
        if (!match.Success) return null;

        // Parse directive
        var directiveType = match.Groups["directive"].Value.ToLowerInvariant();
        var directiveValue = match.Groups["value"].Value.Trim();

        return new SlideDirective(directiveType, directiveValue);
    }
}