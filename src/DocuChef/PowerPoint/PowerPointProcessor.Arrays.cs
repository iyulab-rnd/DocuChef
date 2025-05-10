namespace DocuChef.PowerPoint;

using DocuChef.PowerPoint.Helpers;

/// <summary>
/// Partial class for PowerPointProcessor - Array handling methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Find all array references in a slide's text elements
    /// </summary>
    private List<ArrayReference> FindArrayReferences(SlidePart slidePart)
    {
        var result = new List<ArrayReference>();
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();

        // Detailed logging
        Logger.Debug($"[FIND-ARRAY-REF] Analyzing {shapes.Count} shapes for array references");

        // Regular expression to find array indices like ${Items[0].Name} or Items[3]
        var dollarSignRegex = new Regex(@"\${(\w+)\[(\d+)\](\.[\w]+)?}");
        var directRegex = new Regex(@"(\w+)\[(\d+)\](\.[\w]+)?");

        // New pattern for functions with array parameters like ${ppt.Image(Items[0].ImageUrl)}
        var functionRegex = new Regex(@"\${ppt\.(\w+)\((\w+)\[(\d+)\][^)]*\)}");

        foreach (var shape in shapes)
        {
            var textRuns = shape.Descendants<A.Text>().ToList();
            string shapeName = shape.GetShapeName();

            if (textRuns.Count > 0)
            {
                Logger.Debug($"[FIND-ARRAY-REF] Checking shape '{shapeName ?? "(unnamed)"}' with {textRuns.Count} text runs");
            }

            foreach (var textRun in textRuns)
            {
                if (string.IsNullOrEmpty(textRun.Text))
                    continue;

                // Log the text being analyzed
                Logger.Debug($"[FIND-ARRAY-REF] Analyzing text: {textRun.Text}");

                // Check for ${array[index].property} pattern
                var dollarMatches = dollarSignRegex.Matches(textRun.Text);
                if (dollarMatches.Count > 0)
                {
                    Logger.Debug($"[FIND-ARRAY-REF] Found {dollarMatches.Count} direct array references");
                }

                foreach (Match match in dollarMatches)
                {
                    if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int index))
                    {
                        string arrayName = match.Groups[1].Value;
                        string propPath = match.Groups[3].Success ? match.Groups[3].Value : "";

                        Logger.Debug($"[FIND-ARRAY-REF] Found direct reference: {arrayName}[{index}]{propPath}");

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
                if (directMatches.Count > 0)
                {
                    Logger.Debug($"[FIND-ARRAY-REF] Found {directMatches.Count} plain array references");
                }

                foreach (Match match in directMatches)
                {
                    if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int index))
                    {
                        string arrayName = match.Groups[1].Value;
                        string propPath = match.Groups[3].Success ? match.Groups[3].Value : "";

                        Logger.Debug($"[FIND-ARRAY-REF] Found plain reference: {arrayName}[{index}]{propPath}");

                        result.Add(new ArrayReference
                        {
                            ArrayName = arrayName,
                            Index = index,
                            PropertyPath = propPath,
                            Pattern = match.Value
                        });
                    }
                }

                // Check for function with array parameters pattern
                var functionMatches = functionRegex.Matches(textRun.Text);
                if (functionMatches.Count > 0)
                {
                    Logger.Debug($"[FIND-ARRAY-REF] Found {functionMatches.Count} function array references");
                }

                foreach (Match match in functionMatches)
                {
                    if (match.Groups.Count > 3 && int.TryParse(match.Groups[3].Value, out int index))
                    {
                        string functionName = match.Groups[1].Value;
                        string arrayName = match.Groups[2].Value;

                        Logger.Debug($"[FIND-ARRAY-REF] Found function reference: {functionName}({arrayName}[{index}]...)");

                        result.Add(new ArrayReference
                        {
                            ArrayName = arrayName,
                            Index = index,
                            PropertyPath = "",
                            Pattern = match.Value
                        });
                    }
                }
            }
        }

        // Log summary
        if (result.Count > 0)
        {
            var arrayGroups = result.GroupBy(r => r.ArrayName);
            foreach (var group in arrayGroups)
            {
                Logger.Debug($"[FIND-ARRAY-REF] Found {group.Count()} references to array '{group.Key}' with max index {group.Max(r => r.Index)}");
            }
        }
        else
        {
            Logger.Debug("[FIND-ARRAY-REF] No array references found in this slide");
        }

        return result;
    }

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

            // Add safety check
            if (itemsPerSlide > 100)
            {
                Logger.Warning($"[ANALYZE-DEBUG] Unusually high items per slide: {itemsPerSlide} - verify calculations");
                // Consider imposing a reasonable limit
                itemsPerSlide = Math.Min(itemsPerSlide, 30); // Cap at a sane number
                Logger.Debug($"[ANALYZE-DEBUG] Capped items per slide to {itemsPerSlide}");
            }

            // Get array from variables
            object arrayObj = ResolveVariableValue(arrayName);
            if (arrayObj == null)
            {
                Logger.Warning($"Array '{arrayName}' not found in variables");
                continue;
            }

            // Convert to list for processing
            var items = ConvertToList(arrayObj);
            if (items == null)
            {
                Logger.Warning($"Array '{arrayName}' could not be converted to list");
                continue;
            }

            // If array has fewer or equal items than max index + 1, no duplication needed
            if (items.Count <= itemsPerSlide)
            {
                Logger.Debug($"Array '{arrayName}' has {items.Count} items, no duplication needed for {itemsPerSlide} items per slide");
                continue;
            }

            // Calculate needed slides
            int slidesNeeded = (int)Math.Ceiling((double)items.Count / itemsPerSlide);
            slidesNeeded = Math.Min(slidesNeeded, _options.MaxSlidesFromTemplate);

            Logger.Info($"Array '{arrayName}' has {items.Count} items, needs {slidesNeeded} slides with {itemsPerSlide} items per slide");

            // Duplicate slides and process each with its data batch
            DuplicateTemplateSlides(presentationPart, slidePart, arrayName, items, itemsPerSlide, slidesNeeded, slideIndex);
        }
    }

    /// <summary>
    /// Duplicate template slides first, then process each slide with its batch of data
    /// </summary>
    private void DuplicateTemplateSlides(
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

        // Add detailed logging
        Logger.Debug($"====== ARRAY BATCH DUPLICATION DETAILS ======");
        Logger.Debug($"Array name: {arrayName}");
        Logger.Debug($"Total items: {items.Count}");
        Logger.Debug($"Items per slide: {itemsPerSlide}");
        Logger.Debug($"Slides needed: {slidesNeeded}");
        Logger.Debug($"Template index: {templateIndex}");
        Logger.Debug($"Original slide index in presentation: {originalIndex}");
        Logger.Debug($"Maximum slide ID: {maxSlideId}");
        Logger.Debug($"===============================");

        // Store original template path to avoid reprocessing later
        string templateRelId = presentationPart.GetIdOfPart(templateSlidePart);
        HashSet<string> processedRelIds = new HashSet<string> { templateRelId };

        // Create a list to hold all the slide parts (original + duplicates)
        List<SlidePart> allSlideParts = new List<SlidePart> { templateSlidePart };

        // First phase: Create all the duplicates and update array references
        for (int i = 1; i < slidesNeeded; i++)
        {
            try
            {
                // Calculate batch start index for this slide
                int batchStartIndex = i * itemsPerSlide;

                // Add safety check
                if (batchStartIndex >= items.Count)
                {
                    Logger.Debug($"Batch start index {batchStartIndex} exceeds items count {items.Count}, skipping additional slide creation");
                    break;
                }

                // Double-check for reasonable index
                if (batchStartIndex > 100)
                {
                    Logger.Warning($"[DUPLICATE-DEBUG] Unusually high batch start index: {batchStartIndex} - verify calculation");
                }

                Logger.Debug($"Creating duplicate slide {i + 1} with batch start index: {batchStartIndex}");

                // Clone the original template slide and update array references
                SlidePart newSlidePart = CloneSlideForArrayBatch(
                    presentationPart,
                    templateSlidePart,
                    arrayName,
                    batchStartIndex,
                    itemsPerSlide);

                string newRelId = presentationPart.GetIdOfPart(newSlidePart);
                processedRelIds.Add(newRelId);

                // Add new slide to presentation immediately after the previous one
                var newSlideId = new SlideId
                {
                    Id = maxSlideId + (uint)i,
                    RelationshipId = newRelId
                };

                slideIdList.InsertAt(newSlideId, insertPosition++);
                allSlideParts.Add(newSlidePart);

                Logger.Debug($"Created duplicate slide {i + 1} with relationship ID {newRelId} for batch starting at {batchStartIndex}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error creating duplicate slide {i + 1}: {ex.Message}", ex);
                // Continue with other slides even if one fails
            }
        }

        // Save the presentation after adding all slides
        presentationPart.Presentation.Save();

        // Second phase: Process each slide with its corresponding data batch
        for (int slideIndex = 0; slideIndex < allSlideParts.Count; slideIndex++)
        {
            try
            {
                int batchStartIndex = slideIndex * itemsPerSlide;

                // Double-check for reasonable index
                if (batchStartIndex > 100)
                {
                    Logger.Warning($"[DUPLICATE-DEBUG] Unusually high batch start index: {batchStartIndex} for processing - verify calculation");
                }

                // Skip if this batch would be beyond our data range
                if (batchStartIndex >= items.Count)
                {
                    Logger.Debug($"Batch start index {batchStartIndex} exceeds items count {items.Count}, skipping slide processing");
                    continue;
                }

                SlidePart slidePart = allSlideParts[slideIndex];
                string relId = presentationPart.GetIdOfPart(slidePart);

                Logger.Debug($"Processing slide {slideIndex + 1} with data batch starting at index {batchStartIndex}");

                // Create an independent copy of the items list with limited count for safety
                var availableItemCount = Math.Min(items.Count - batchStartIndex, itemsPerSlide);
                Logger.Debug($"Processing batch with {availableItemCount} available items (out of {itemsPerSlide} max per slide)");

                // Additional validation before processing
                if (batchStartIndex < 0 || batchStartIndex >= items.Count)
                {
                    Logger.Warning($"[BATCH-DEBUG] Invalid batch start index: {batchStartIndex}. Skipping this slide.");
                    continue;
                }

                // Pass explicit batch range parameters
                ProcessSlideWithArrayBatch(slidePart, arrayName, items, batchStartIndex, itemsPerSlide);

                // Save each processed slide
                slidePart.Slide.Save();

                // Mark this slide as being processed with array batch data
                _context.ProcessedArraySlides.Add(relId);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing slide {slideIndex + 1}: {ex.Message}", ex);
            }
        }

        Logger.Debug($"Completed processing all slides for array '{arrayName}'");
    }

    /// <summary>
    /// Clone and update slide for a specific batch of array data
    /// </summary>
    private SlidePart CloneSlideForArrayBatch(
        PresentationPart presentationPart,
        SlidePart sourceSlidePart,
        string arrayName,
        int batchStartIndex,
        int itemsPerSlide)
    {
        // First, clone the slide with all its relationships
        SlidePart newSlidePart = CloneSlideWithRelationships(presentationPart, sourceSlidePart);

        Logger.Debug($"Updating array references in cloned slide for batch starting at {batchStartIndex}");

        // Now update array references in all shape texts
        var shapes = newSlidePart.Slide.Descendants<P.Shape>().Where(s => s.TextBody != null).ToList();
        Logger.Debug($"Found {shapes.Count} shapes with text body to update references");

        // Track which shapes we've already processed to avoid double-processing
        HashSet<string> processedShapeIds = new HashSet<string>();

        foreach (var shape in shapes)
        {
            string shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString() ?? "";
            string shapeName = shape.GetShapeName();
            Logger.Debug($"Checking shape: {shapeName ?? "(unnamed)"} (ID: {shapeId})");

            // Skip if already processed
            if (!string.IsNullOrEmpty(shapeId) && processedShapeIds.Contains(shapeId))
            {
                Logger.Debug($"Skipping already processed shape: {shapeName ?? "(unnamed)"}");
                continue;
            }

            // Get all text runs
            var textRuns = shape.Descendants<A.Text>().ToList();
            Logger.Debug($"Found {textRuns.Count} text runs in shape");

            foreach (var textRun in textRuns)
            {
                if (string.IsNullOrEmpty(textRun.Text))
                    continue;

                string originalText = textRun.Text;
                Logger.Debug($"Original text: {originalText}");

                // Replace array references in text including inside function calls
                string updatedText = UpdateArrayReferencesWithFunctionSupport(originalText, arrayName, batchStartIndex);

                if (updatedText != originalText)
                {
                    textRun.Text = updatedText;
                    Logger.Debug($"Updated text from '{originalText}' to '{updatedText}'");
                }
                else
                {
                    Logger.Debug($"No updates needed for text: {originalText}");
                }
            }

            // Mark as processed
            if (!string.IsNullOrEmpty(shapeId))
            {
                processedShapeIds.Add(shapeId);
            }
        }

        // Save the updated slide
        newSlidePart.Slide.Save();

        return newSlidePart;
    }

    /// <summary>
    /// Get complete text content from a shape
    /// </summary>
    private string GetShapeTextContent(P.Shape shape)
    {
        if (shape?.TextBody == null)
            return string.Empty;

        var sb = new StringBuilder();

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            if (sb.Length > 0)
                sb.AppendLine();

            foreach (var run in paragraph.Elements<A.Run>())
            {
                var text = run.GetFirstChild<A.Text>();
                if (text != null && !string.IsNullOrEmpty(text.Text))
                {
                    sb.Append(text.Text);
                }
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Update array references in text for a specific batch with support for function calls
    /// </summary>
    private string UpdateArrayReferencesWithFunctionSupport(string text, string arrayName, int batchStartIndex)
    {
        if (string.IsNullOrEmpty(text) || batchStartIndex == 0)
            return text;

        // 변경 전 로그
        Logger.Debug($"[ARRAY-DEBUG] Updating array references in text: '{text}', startIndex={batchStartIndex}");

        // 이미 업데이트된 텍스트인지 확인 (이미 높은 인덱스가 있는 경우)
        if (IsTextAlreadyUpdated(text, arrayName, batchStartIndex))
        {
            Logger.Debug($"[ARRAY-DEBUG] Text already contains updated indices, skipping: '{text}'");
            return text;
        }

        // Pattern 1: ${arrayName[index]} - Direct array reference in expression
        var result = Regex.Replace(text, $"\\${{{arrayName}\\[(\\d+)\\]", match =>
        {
            if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int localIndex))
            {
                // *** 인덱스 계산 디버깅 ***
                Logger.Debug($"[ARRAY-DEBUG] Direct Reference: localIndex={localIndex}, batchStartIndex={batchStartIndex}");

                // 명시적 범위 검사 추가
                if (localIndex > 100 || localIndex < 0)
                {
                    Logger.Warning($"[ARRAY-DEBUG] Invalid local index: {localIndex} - out of reasonable range");
                    return match.Value; // 유효하지 않은 인덱스는 변경하지 않음
                }

                // Calculate new global index (with safety check)
                int newIndex = batchStartIndex + localIndex;

                // 합산 결과 검증
                if (newIndex > 100 || newIndex < 0)
                {
                    Logger.Warning($"[ARRAY-DEBUG] Invalid calculated index: {newIndex} - out of reasonable range");
                    return match.Value; // 유효하지 않은 결과는 변경하지 않음
                }

                // Replace with new index
                string replacement = $"${{{arrayName}[{newIndex}]";
                Logger.Debug($"[ARRAY-DEBUG] Replaced direct array reference: {match.Value} -> {replacement}");
                return replacement;
            }
            return match.Value;
        });

        // Pattern 2: ${ppt.Image(arrayName[index])} - Array reference in function call
        // 중요: 여기서는 하나의 정규식만 사용하여 일관되게 처리하도록 개선
        result = Regex.Replace(result, $"\\${{ppt\\.(\\w+)\\({arrayName}\\[(\\d+)\\](\\.[\\w]+)?\\)}}", match =>
        {
            if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int localIndex))
            {
                // *** 인덱스 계산 디버깅 ***
                string functionName = match.Groups[1].Value;
                string propertyPath = match.Groups.Count > 3 ? match.Groups[3].Value : "";

                Logger.Debug($"[ARRAY-DEBUG] Function Reference: functionName={functionName}, localIndex={localIndex}, batchStartIndex={batchStartIndex}, propertyPath={propertyPath}");

                // 명시적 범위 검사 추가
                if (localIndex > 100 || localIndex < 0)
                {
                    Logger.Warning($"[ARRAY-DEBUG] Invalid local index in function: {localIndex} - out of reasonable range");
                    return match.Value; // 유효하지 않은 인덱스는 변경하지 않음
                }

                int newIndex = batchStartIndex + localIndex;

                // 합산 결과 검증
                if (newIndex > 100 || newIndex < 0)
                {
                    Logger.Warning($"[ARRAY-DEBUG] Invalid calculated index in function: {newIndex} - out of reasonable range");
                    return match.Value; // 유효하지 않은 결과는 변경하지 않음
                }

                string replacement = $"${{ppt.{functionName}({arrayName}[{newIndex}]{propertyPath})}}";
                Logger.Debug($"[ARRAY-DEBUG] Replaced function array reference: {match.Value} -> {replacement}");
                return replacement;
            }
            return match.Value;
        });

        // 변경 후 로그
        if (result != text)
        {
            Logger.Debug($"[ARRAY-DEBUG] Text after all replacements: '{result}'");
        }
        else
        {
            Logger.Debug($"[ARRAY-DEBUG] No array references updated in text: '{text}'");
        }

        return result;
    }

    private bool IsTextAlreadyUpdated(string text, string arrayName, int batchStartIndex)
    {
        // Skip updating if text already contains indices higher than the batch start
        var highIndexPattern = $"{arrayName}\\[(\\d+)\\]";
        var matches = Regex.Matches(text, highIndexPattern);

        // Check if any index is already higher than what we'd normally expect
        foreach (Match match in matches)
        {
            if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
            {
                // If index is significantly higher than batch start, likely already processed
                if (index >= batchStartIndex && index > 10) // 10 is arbitrary threshold
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Process slide with array batch, ensuring proper variable replacement
    /// </summary>
    private void ProcessSlideWithArrayBatch(SlidePart slidePart, string arrayName, List<object> items, int startIndex, int itemsPerSlide)
    {
        Logger.Debug($"Processing slide with {items.Count} items, starting at index {startIndex} with {itemsPerSlide} items per slide");

        // Update context for this slide
        _context.SlidePart = slidePart;

        // Create a batch-specific variable dictionary
        var batchVariables = new Dictionary<string, object>(_context.Variables);

        // Add batch metadata
        int batchIndex = startIndex / itemsPerSlide;
        batchVariables["_batchIndex"] = batchIndex;
        batchVariables["_batchStartIndex"] = startIndex;
        batchVariables["_batchEndIndex"] = Math.Min(startIndex + itemsPerSlide - 1, items.Count - 1);
        batchVariables["_batchSize"] = Math.Min(itemsPerSlide, items.Count - startIndex);
        batchVariables["_totalItems"] = items.Count;

        // Important: Calculate the number of available items for this batch
        int availableItemCount = Math.Min(items.Count - startIndex, itemsPerSlide);
        Logger.Debug($"Batch has {availableItemCount} available items out of {itemsPerSlide} possible items per slide");

        // 먼저 모든 도형에 대해 사전 검사 실행
        Logger.Debug("Performing pre-scan of all shapes to identify shapes with array references");
        var outOfRangeShapes = new List<KeyValuePair<P.Shape, int>>();
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();

        // 특정 인덱스 패턴을 가진 도형들 확인
        foreach (var shape in shapes)
        {
            if (shape.TextBody == null)
                continue;

            string text = GetShapeTextContent(shape);
            if (string.IsNullOrEmpty(text))
                continue;

            // 배열 인덱스 찾기
            var indexMatches = Regex.Matches(text, $"{arrayName}\\[(\\d+)\\]");
            foreach (Match match in indexMatches)
            {
                if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
                {
                    // 인덱스가 가용 범위를 벗어나는지 확인
                    if (index >= availableItemCount)
                    {
                        string shapeName = shape.GetShapeName();
                        Logger.Debug($"Shape '{shapeName ?? "(unnamed)"}' has array reference to index {index} which is out of range (available: {availableItemCount})");
                        outOfRangeShapes.Add(new KeyValuePair<P.Shape, int>(shape, index));
                    }
                }
            }
        }

        // 범위를 벗어나는 인덱스를 참조하는 도형 숨기기
        foreach (var shapePair in outOfRangeShapes)
        {
            var shape = shapePair.Key;
            int index = shapePair.Value;
            string shapeName = shape.GetShapeName();

            // Shape를 숨기는 강력한 방법 적용
            HideShapeCompletely(shape);

            Logger.Debug($"Hidden shape '{shapeName ?? "(unnamed)"}' that references out-of-range index {index}");
        }

        // Create a ShapeVisibilityManager and hide shapes with out-of-range references
        var visibilityManager = new ShapeVisibilityManager(_context);
        visibilityManager.HideShapesWithOutOfRangeIndices(slidePart, arrayName, availableItemCount, startIndex, itemsPerSlide);

        // Retrieve and log all shapes after visibility changes
        shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        Logger.Debug($"Found {shapes.Count} shapes in slide after visibility processing");

        // 중요: 먼저 직접 배열 참조 확인 (디버깅용)
        foreach (var shape in shapes)
        {
            if (shape.TextBody == null)
                continue;

            // 이 도형이 숨김 처리되었는지 확인
            bool isHidden = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value ?? false;
            if (isHidden)
            {
                string shapeName = shape.GetShapeName();
                Logger.Debug($"Skipping hidden shape: {shapeName ?? "(unnamed)"}");
                continue;
            }

            string text = GetShapeTextContent(shape);
            if (string.IsNullOrEmpty(text))
                continue;

            Logger.Debug($"Shape text: {text}");

            // 배열 참조가 들어있는지 확인
            if (text.Contains($"{arrayName}["))
            {
                string shapeName = shape.GetShapeName();
                Logger.Debug($"Shape '{shapeName ?? "(unnamed)"}' contains array reference: {text}");

                // 특히 이미지 함수 내 배열 참조 확인
                if (text.Contains("ppt.Image(") && text.Contains($"{arrayName}["))
                {
                    Logger.Debug($"*** IMPORTANT *** Found image function with array reference: {text}");

                    // 배열 인덱스 추출 
                    var match = Regex.Match(text, $"{arrayName}\\[(\\d+)\\]");
                    if (match.Success)
                    {
                        int index = int.Parse(match.Groups[1].Value);
                        Logger.Debug($"Array index in image function: {index}, batch start index: {startIndex}");

                        // 인덱스가 올바른지 확인
                        if (startIndex > 0 && index < startIndex)
                        {
                            Logger.Warning($"*** WRONG INDEX *** Image function has wrong index: {index}, expected >= {startIndex}");
                        }

                        // 인덱스가 가용 데이터 범위를 벗어나는지 확인
                        if (index >= availableItemCount + startIndex)
                        {
                            Logger.Warning($"*** OUT OF RANGE *** Image function has index beyond available data: {index}, available range: {startIndex} to {startIndex + availableItemCount - 1}");

                            // 즉시 이 도형 숨기기
                            HideShapeCompletely(shape);
                            continue;
                        }
                    }
                }
            }
        }

        // 먼저 이미지 함수가 포함된 도형을 처리
        foreach (var shape in shapes)
        {
            // Skip hidden shapes immediately
            bool isHidden = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value ?? false;
            if (isHidden)
            {
                continue;
            }

            string shapeName = shape.GetShapeName();

            // Skip shapes without text body
            if (shape.TextBody == null)
                continue;

            string text = GetShapeTextContent(shape);

            // 이미지 함수 패턴 확인
            if (!string.IsNullOrEmpty(text) && (
                text.Contains("${ppt.Image(") ||
                text.Contains("ppt.Image("))
            )
            {
                Logger.Debug($"Processing image function in shape: {shapeName ?? "(unnamed)"}, text: {text}");

                // 형식에 따라 다르게 처리
                if (text.Contains("${ppt.Image("))
                {
                    // 표현식으로 감싸진 이미지 함수
                    var match = Regex.Match(text, @"\${ppt\.Image\(([^)]+)\)}");
                    if (match.Success)
                    {
                        string param = match.Groups[1].Value;
                        Logger.Debug($"Image function parameter: {param}");

                        // 배열 참조인지 확인
                        if (param.Contains($"{arrayName}["))
                        {
                            var indexMatch = Regex.Match(param, $"{arrayName}\\[(\\d+)\\]");
                            if (indexMatch.Success)
                            {
                                int index = int.Parse(indexMatch.Groups[1].Value);

                                // 올바른 배치 인덱스인지 확인
                                if (startIndex > 0 && (index < startIndex || index >= startIndex + itemsPerSlide))
                                {
                                    Logger.Warning($"*** INCORRECT INDEX *** Found index {index} outside batch range [{startIndex}-{startIndex + itemsPerSlide - 1}]");
                                }

                                // 가용 데이터 범위를 벗어나는지 확인
                                if (index >= startIndex + availableItemCount)
                                {
                                    Logger.Warning($"*** OUT OF RANGE *** Found index {index} outside available data range [{startIndex}-{startIndex + availableItemCount - 1}]");

                                    // 이 도형은 가용 데이터 범위를 벗어나므로 숨김 처리하고 건너뛰기
                                    HideShapeCompletely(shape);
                                    continue;
                                }
                            }
                        }
                    }
                }

                // Update context
                UpdateShapeContext(shape);

                // Process PowerPoint functions
                bool processed = ProcessPowerPointFunctions(shape);

                if (processed)
                {
                    Logger.Debug($"Successfully processed image function in shape: {shapeName ?? "(unnamed)"}");
                }
                else
                {
                    Logger.Warning($"Failed to process image function in shape: {shapeName ?? "(unnamed)"}");
                }
            }
        }

        // Process each shape for text content
        foreach (var shape in shapes)
        {
            // Skip hidden shapes immediately
            bool isHidden = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value ?? false;
            if (isHidden)
            {
                continue;
            }

            string shapeName = shape.GetShapeName();
            Logger.Debug($"Processing array batch shape: {shapeName ?? "(unnamed)"}");

            // Skip shapes without text body
            if (shape.TextBody == null)
                continue;

            // Try to process with FormattedTextProcessor 
            var formattedTextProcessor = new FormattedTextProcessor(this, batchVariables);
            bool processed = formattedTextProcessor.ProcessShapeTextWithFormatting(shape);

            // If that didn't work, try other methods
            if (!processed)
            {
                // Only process shapes that contain array reference expressions
                bool containsArrayRefs = ContainsArrayReference(shape, arrayName);
                if (containsArrayRefs)
                {
                    string text = shape.GetText();
                    if (!string.IsNullOrEmpty(text))
                    {
                        // Check if the text contains out-of-range indices
                        bool hasOutOfRangeIndices = HasOutOfRangeArrayIndices(text, arrayName, availableItemCount);
                        if (hasOutOfRangeIndices)
                        {
                            // 범위를 벗어나는 인덱스가 있는 도형은 숨김 처리
                            HideShapeCompletely(shape);
                            Logger.Debug($"Hidden shape '{shapeName ?? "(unnamed)"}' containing out-of-range indices");
                            continue;
                        }

                        // Process text and update shape if it changed
                        string processedText = ProcessArrayTextExpressions(text, arrayName, batchVariables);
                        if (processedText != text)
                        {
                            shape.SetText(processedText);
                        }
                    }
                }
            }
        }

        // Process any PowerPoint functions in the slide
        ProcessPowerPointFunctionsInSlide(slidePart);

        // 마지막으로 슬라이드 검사하여 아직 숨겨지지 않은 범위 외 인덱스 참조를 가진 도형 찾기
        var remainingShapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        foreach (var shape in remainingShapes)
        {
            // 이미 숨겨진 도형은 건너뛰기
            bool isHidden = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value ?? false;
            if (isHidden)
                continue;

            if (shape.TextBody == null)
                continue;

            string text = GetShapeTextContent(shape);
            if (string.IsNullOrEmpty(text))
                continue;

            if (HasOutOfRangeArrayIndices(text, arrayName, availableItemCount))
            {
                string shapeName = shape.GetShapeName();
                Logger.Warning($"Found remaining shape '{shapeName ?? "(unnamed)"}' with out-of-range indices - hiding it now");
                HideShapeCompletely(shape);
            }
        }

        // Save the slide
        slidePart.Slide.Save();
    }

    /// <summary>
    /// 도형을 완전히 숨기는 강화된 메서드
    /// </summary>
    private void HideShapeCompletely(P.Shape shape)
    {
        if (shape == null)
            return;

        try
        {
            // 1. Hidden 속성 설정
            var nvProps = shape.NonVisualShapeProperties;
            if (nvProps != null)
            {
                var nvDrawProps = nvProps.NonVisualDrawingProperties;
                if (nvDrawProps != null)
                {
                    // 속성이 없으면 새로 생성합니다
                    if (nvDrawProps.Hidden == null)
                    {
                        nvDrawProps.Hidden = new BooleanValue(true);
                    }
                    else
                    {
                        nvDrawProps.Hidden.Value = true;
                    }
                }
            }

            // 2. 크기를 최소화
            if (shape.ShapeProperties?.Transform2D?.Extents != null)
            {
                var extents = shape.ShapeProperties.Transform2D.Extents;
                var nvAppProps = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;

                if (nvAppProps != null)
                {
                    long cx = extents.Cx?.Value ?? 0;
                    long cy = extents.Cy?.Value ?? 0;

                    // 원래 크기 저장
                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originalcx", "", cx.ToString()));
                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originalcy", "", cy.ToString()));

                    // 크기를 1로 설정
                    extents.Cx = 1;
                    extents.Cy = 1;
                }
            }

            // 3. 위치를 슬라이드 밖으로 이동
            if (shape.ShapeProperties?.Transform2D?.Offset != null)
            {
                var offset = shape.ShapeProperties.Transform2D.Offset;
                var nvAppProps = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;

                if (nvAppProps != null)
                {
                    long x = offset.X?.Value ?? 0;
                    long y = offset.Y?.Value ?? 0;

                    // 원래 위치 저장
                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originalx", "", x.ToString()));
                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originaly", "", y.ToString()));

                    // 위치를 슬라이드 밖으로 이동 (-10000000, -10000000)
                    offset.X = -10000000;
                    offset.Y = -10000000;
                }
            }

            // 4. 투명도를 100%로 설정
            if (shape.ShapeProperties != null)
            {
                var fillProps = shape.ShapeProperties.GetFirstChild<A.SolidFill>();
                if (fillProps == null)
                {
                    fillProps = new A.SolidFill();
                    shape.ShapeProperties.AppendChild(fillProps);
                }

                var transparency = fillProps.GetFirstChild<A.Alpha>();
                if (transparency == null)
                {
                    transparency = new A.Alpha() { Val = 0 }; // 0% 불투명 = 100% 투명
                    fillProps.AppendChild(transparency);
                }
                else
                {
                    transparency.Val = 0;
                }
            }

            // 5. 텍스트를 지움
            if (shape.TextBody != null)
            {
                shape.ClearText();
            }

            string shapeName = shape.GetShapeName();
            Logger.Debug($"Completely hidden shape: '{shapeName ?? "(unnamed)"}'");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error hiding shape completely: {ex.Message}");
        }
    }

    /// <summary>
    /// 텍스트에 사용 가능한 항목 수를 초과하는 배열 인덱스가 포함되어 있는지 확인
    /// </summary>
    private bool HasOutOfRangeArrayIndices(string text, string arrayName, int availableItemCount)
    {
        if (string.IsNullOrEmpty(text) || !text.Contains($"{arrayName}["))
            return false;

        // 모든 배열 참조 추출
        var matches = Regex.Matches(text, $"{arrayName}\\[(\\d+)\\]");
        foreach (Match match in matches)
        {
            if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
            {
                if (index >= availableItemCount)
                {
                    Logger.Debug($"Found out-of-range index {index} in text (available: {availableItemCount})");
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Check if shape contains array references to the specified array
    /// </summary>
    private bool ContainsArrayReference(P.Shape shape, string arrayName)
    {
        if (shape.TextBody == null)
            return false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var text = run.GetFirstChild<A.Text>();
                if (text != null && !string.IsNullOrEmpty(text.Text))
                {
                    if (text.Text.Contains($"${{{arrayName}[") || text.Text.Contains($"{arrayName}["))
                        return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Process array references in text
    /// </summary>
    private string ProcessArrayTextExpressions(string text, string arrayName, Dictionary<string, object> variables)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Process ${Array[index].Property} expressions
        var pattern = $"\\${{{arrayName}\\[(\\d+)\\](\\.\\w+)?(:[^}}]+)?}}";
        return Regex.Replace(text, pattern, match =>
        {
            int index = int.Parse(match.Groups[1].Value);
            string propPath = match.Groups[2].Success ? match.Groups[2].Value : null;
            string format = match.Groups[3].Success ? match.Groups[3].Value : null;

            // Build variable key
            string key = propPath != null
                ? $"{arrayName}[{index}]{propPath}"
                : $"{arrayName}[{index}]";

            // Check if we have this key in our variables
            if (variables.TryGetValue(key, out var value) && value != null)
            {
                // Apply formatting if specified
                if (!string.IsNullOrEmpty(format) && format.StartsWith(":") && value is IFormattable formattable)
                {
                    return formattable.ToString(format.Substring(1), System.Globalization.CultureInfo.CurrentCulture);
                }

                return value.ToString();
            }

            // If value not found or null, keep the original expression
            return match.Value;
        });
    }
}