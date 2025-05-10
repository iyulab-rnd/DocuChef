using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Slide handling methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Process a single slide
    /// </summary>
    private void ProcessSlide(PresentationPart presentationPart, SlideId slideId, int slideIndex)
    {
        Logger.Debug($"Processing slide {slideIndex} with ID {slideId.RelationshipId}");

        // Update slide context
        _context.Slide.Index = slideIndex;
        _context.Slide.Id = slideId.RelationshipId;

        // Get slide part
        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
        if (slidePart == null || slidePart.Slide == null)
        {
            Logger.Warning($"Slide part not found for ID {slideId.RelationshipId}");
            return;
        }

        // Store SlidePart in context
        _context.SlidePart = slidePart;

        // Get slide notes
        string slideNotes = slidePart.GetNotes();
        _context.Slide.Notes = slideNotes;

        Logger.Debug($"Slide notes: {slideNotes}");

        // Parse directives from slide notes using enhanced DirectiveParser
        var directives = DirectiveParser.ParseDirectives(slideNotes);

        // Process directives (e.g., #if)
        foreach (var directive in directives)
        {
            Logger.Debug($"Processing directive: {directive.Name}");
            ProcessShapeDirective(slidePart, directive);
        }

        // Analyze array references and handle automatic slide duplication
        AnalyzeSlideForArrayIndices(presentationPart, slidePart, slideIndex);

        // Process text replacements using DollarSignEngine
        Logger.Debug($"Processing text replacements with DollarSignEngine on slide {slideIndex}");
        ProcessTextReplacements(slidePart);

        // Process special PowerPoint functions
        ProcessPowerPointFunctionsInSlide(slidePart);

        // Save slide after processing
        try
        {
            slidePart.Slide.Save();
            Logger.Debug($"Slide {slideIndex} saved after processing");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error saving slide {slideIndex}: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Process PowerPoint functions in slide
    /// </summary>
    private void ProcessPowerPointFunctionsInSlide(SlidePart slidePart)
    {
        try
        {
            var shapes = slidePart.Slide.Descendants<Shape>().ToList();
            Logger.Debug($"Processing PowerPoint functions in {shapes.Count} shapes");

            bool hasChanges = false;

            foreach (var shape in shapes)
            {
                // Update shape context
                _context.Shape = new ShapeContext
                {
                    Name = shape.GetShapeName(),
                    Id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
                    Text = shape.GetText(),
                    ShapeObject = shape
                };

                // Process PowerPoint functions
                bool shapeChanged = ProcessPowerPointFunctions(shape);
                if (shapeChanged)
                    hasChanges = true;
            }

            // Save if any changes were made
            if (hasChanges)
            {
                slidePart.Slide.Save();
                Logger.Debug("Slide saved after processing PowerPoint functions");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing PowerPoint functions: {ex.Message}", ex);
        }
    }

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
}