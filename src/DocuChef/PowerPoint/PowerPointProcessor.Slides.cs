using DocuChef.PowerPoint.Helpers;

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
        string relationshipId = slideId.RelationshipId;
        Logger.Debug($"Processing slide {slideIndex} with ID {relationshipId}");

        // Check if this slide has already been processed with array batch data
        if (_context.ProcessedArraySlides.Contains(relationshipId))
        {
            Logger.Debug($"Skipping slide {slideIndex} as it was already processed with array batch data");
            return;
        }

        // Update slide context
        _context.Slide.Index = slideIndex;
        _context.Slide.Id = relationshipId;

        // Get slide part
        var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
        if (slidePart == null || slidePart.Slide == null)
        {
            Logger.Warning($"Slide part not found for ID {relationshipId}");
            return;
        }

        // Store SlidePart in context
        _context.SlidePart = slidePart;

        // Get slide notes
        string slideNotes = slidePart.GetNotes();
        _context.Slide.Notes = slideNotes;

        Logger.Debug($"Slide notes: {slideNotes}");

        // 모든 배열 변수의 길이 수집
        var arrayLengths = CollectArrayLengths();

        // 범위를 벗어나는 도형 선제적으로 숨김 처리
        var outOfRangeHandler = new OutOfRangeShapeHandler(_context);
        outOfRangeHandler.ScanAndHideOutOfRangeShapes(slidePart, arrayLengths);

        // 숨겨진 도형이 있을 경우 로그 출력
        int hiddenCount = outOfRangeHandler.GetHiddenShapeCount();
        if (hiddenCount > 0)
        {
            Logger.Info($"Hidden {hiddenCount} shapes with out-of-range array indices on slide {slideIndex}");
        }

        // Parse directives from slide notes using enhanced DirectiveParser
        var directives = DirectiveParser.ParseDirectives(slideNotes);

        // Process directives (e.g., #if)
        foreach (var directive in directives)
        {
            Logger.Debug($"Processing directive: {directive.Name}");
            ProcessShapeDirective(slidePart, directive);
        }

        // Analyze array references and handle automatic slide duplication
        // This will mark duplicated slides as processed in _context.ProcessedArraySlides
        AnalyzeSlideForArrayIndices(presentationPart, slidePart, slideIndex);

        // Skip further processing if this slide was processed as part of array duplication
        // This check is needed in case the slide was processed during AnalyzeSlideForArrayIndices
        if (_context.ProcessedArraySlides.Contains(relationshipId))
        {
            Logger.Debug($"Skipping remaining processing for slide {slideIndex} as it was handled by array batch processing");
            return;
        }

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
    /// 모든 배열 변수의 길이 수집
    /// </summary>
    private Dictionary<string, int> CollectArrayLengths()
    {
        var arrayLengths = new Dictionary<string, int>();

        foreach (var entry in _context.Variables)
        {
            if (entry.Value is IList list)
            {
                arrayLengths[entry.Key] = list.Count;
            }
            else if (entry.Value is IEnumerable enumerable && !(entry.Value is string))
            {
                // 컬렉션 길이 계산
                int count = 0;
                foreach (var _ in enumerable)
                {
                    count++;
                }
                arrayLengths[entry.Key] = count;
            }
        }

        // 로그로 배열 길이 표시
        if (arrayLengths.Any())
        {
            foreach (var entry in arrayLengths)
            {
                Logger.Debug($"Array {entry.Key} has {entry.Value} items");
            }
        }

        return arrayLengths;
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

        // 중요: 이 부분을 추가하여 복제된 슬라이드에서도 ppt.Image 함수를 인식하도록 함
        // Process image placeholders in slide (새로 추가된 부분)
        ProcessImagePlaceholders(newSlidePart);

        // Save the slide
        newSlidePart.Slide.Save();

        return newSlidePart;
    }

    private void ProcessImagePlaceholders(SlidePart slidePart)
    {
        try
        {
            var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
            Logger.Debug($"Processing image placeholders in {shapes.Count} shapes in cloned slide");

            foreach (var shape in shapes)
            {
                // Skip shapes without text body
                if (shape.TextBody == null)
                    continue;

                // Check for ppt.Image pattern in text
                foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
                {
                    foreach (var run in paragraph.Elements<A.Run>())
                    {
                        var textElement = run.GetFirstChild<A.Text>();
                        if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                            continue;

                        string text = textElement.Text;
                        if (text.Contains("${ppt.Image(") && text.Contains(")}"))
                        {
                            // Mark this shape for image processing later
                            // We could add a custom attribute or simply preserve the text
                            // as the actual processing will happen during the batch processing phase
                            Logger.Debug($"Marked shape '{shape.GetShapeName() ?? "(unnamed)"}' for image processing");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing image placeholders: {ex.Message}", ex);
        }
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
}