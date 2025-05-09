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
    /// Process a slide-level directive (e.g., #foreach)
    /// </summary>
    private void ProcessSlideDirective(PresentationPart presentationPart, SlidePart slidePart, DirectiveContext directive)
    {
        _context.Directive = directive;

        // Handle foreach directive for slide duplication
        if (directive.Name == "foreach")
        {
            ProcessForeachDirective(presentationPart, slidePart, directive);
        }

        // Add support for other slide-level directives as needed
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
    /// Process foreach directive to create multiple slides from a template slide
    /// </summary>
    private void ProcessForeachDirective(PresentationPart presentationPart, SlidePart templateSlidePart, DirectiveContext directive)
    {
        string collectionName = directive.Value;
        string itemName = directive.Parameters.TryGetValue("itemName", out var itemParam) ? itemParam : "item";

        Logger.Debug($"Processing foreach with collection: {collectionName}, item: {itemName}");

        // Resolve collection from variables
        object collectionObj = ResolveVariableValue(collectionName);
        if (collectionObj == null)
        {
            Logger.Warning($"Collection '{collectionName}' not found for foreach directive");
            return;
        }

        // 컬렉션이 아닌 경우 처리
        if (!(collectionObj is IEnumerable collection))
        {
            Logger.Warning($"Variable '{collectionName}' is not a collection type");
            return;
        }

        // Convert to list for easier processing
        List<object> items = collection.Cast<object>().ToList();
        if (items.Count == 0)
        {
            Logger.Debug("Collection is empty, no slides to generate");
            return;
        }

        // Find the maximum index referenced in slide shapes to determine items per slide
        int maxIndexReferenced = FindMaxIndexReferenced(templateSlidePart, itemName);
        if (maxIndexReferenced < 0)
        {
            Logger.Warning($"No indexed references found for item '{itemName}' in slide, using 1 item per slide");
            maxIndexReferenced = 0; // 기본값 0으로 설정 (1개 항목)
        }

        // Items per slide is maxIndexReferenced + 1 (since indexes are zero-based)
        int itemsPerSlide = maxIndexReferenced + 1;
        Logger.Debug($"Detected {itemsPerSlide} items per slide based on references in template");

        // Calculate number of slides needed
        int slidesNeeded = (int)Math.Ceiling((double)items.Count / itemsPerSlide);
        slidesNeeded = Math.Min(slidesNeeded, _options.MaxSlidesFromTemplate);
        Logger.Debug($"Need to create {slidesNeeded} slides for {items.Count} items");

        // Find template slide in presentation
        var slideIdList = presentationPart.Presentation.SlideIdList;
        var slideIds = slideIdList.ChildElements.OfType<P.SlideId>().ToList();
        int templateIndex = slideIds.FindIndex(id => id.RelationshipId == presentationPart.GetIdOfPart(templateSlidePart));

        if (templateIndex == -1)
        {
            Logger.Error("Template slide not found in presentation");
            return;
        }

        uint maxSlideId = slideIds.Max(id => id.Id.Value);
        int insertPosition = templateIndex + 1;

        // Save original template content
        DocumentFormat.OpenXml.OpenXmlElement originalSlideContent = null;
        try
        {
            // Clone the original slide content for reference
            originalSlideContent = templateSlidePart.Slide.CloneNode(true);
            Logger.Debug("Original slide content saved for reference");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Could not clone original slide content: {ex.Message}");
        }

        // Process each slide
        for (int slideIndex = 0; slideIndex < slidesNeeded; slideIndex++)
        {
            // Calculate item range for this slide
            int startIndex = slideIndex * itemsPerSlide;
            int endIndex = Math.Min(startIndex + itemsPerSlide, items.Count);
            var slideItems = items.Skip(startIndex).Take(endIndex - startIndex).ToList();

            SlidePart slidePart;

            // Use template slide for first batch, clone for others
            if (slideIndex == 0)
            {
                slidePart = templateSlidePart;
            }
            else
            {
                // Clone template slide
                slidePart = CloneSlidePart(presentationPart, templateSlidePart);
                string newRelId = presentationPart.GetIdOfPart(slidePart);

                // Add new slide to presentation
                P.SlideId newSlideId = new P.SlideId
                {
                    Id = maxSlideId + (uint)slideIndex,
                    RelationshipId = newRelId
                };

                slideIdList.InsertAt(newSlideId, insertPosition++);
            }

            // Process this slide with its items
            ProcessSlideItems(slidePart, slideItems, itemName, slideIndex, itemsPerSlide);

            Logger.Debug($"Processed slide {slideIndex + 1} with {slideItems.Count} items");
        }

        // Save presentation
        presentationPart.Presentation.Save();
    }

    /// <summary>
    /// Finds the maximum index referenced in a slide's shapes for a given item name
    /// </summary>
    private int FindMaxIndexReferenced(SlidePart slidePart, string itemName)
    {
        int maxIndex = -1;
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();

        // Regular expression to find item references with indexes: ${item[0].Id} 또는 item[0]
        var regex = new Regex($@"\${{?{itemName}\[(\d+)\](?:\.[\w]+)?}}?");

        foreach (var shape in shapes)
        {
            var textRuns = shape.Descendants<A.Text>().ToList();
            foreach (var textRun in textRuns)
            {
                if (string.IsNullOrEmpty(textRun.Text))
                    continue;

                var matches = regex.Matches(textRun.Text);
                foreach (Match match in matches)
                {
                    if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
                    {
                        maxIndex = Math.Max(maxIndex, index);
                    }
                }
            }
        }

        Logger.Debug($"Maximum index referenced for '{itemName}': {maxIndex}");
        return maxIndex;
    }

    /// <summary>
    /// Process a slide with specific items
    /// </summary>
    private void ProcessSlideItems(SlidePart slidePart, List<object> slideItems, string itemName, int slideIndex, int itemsPerSlide)
    {
        // Update context for this slide
        _context.SlidePart = slidePart;

        // Add batch info to variables
        _context.Variables["_batchIndex"] = slideIndex;
        _context.Variables["_batchStart"] = slideIndex * itemsPerSlide + 1; // 1-based index
        _context.Variables["_batchEnd"] = slideIndex * itemsPerSlide + slideItems.Count;
        _context.Variables["_batchCount"] = slideItems.Count;

        // Create a mapped dictionary for item references
        // This allows template to use ${item[0]}, ${item[1]} etc. which will be mapped to appropriate collection items
        var itemsDict = new Dictionary<string, object>();

        // Populate the dictionary with items AND indexed references for each item
        for (int i = 0; i < itemsPerSlide; i++)
        {
            if (i < slideItems.Count)
            {
                // 전체 항목 배열을 관리하기 위한 특별 항목 추가
                itemsDict[$"{itemName}s"] = slideItems;

                // 개별 항목 참조 설정
                var currentItem = slideItems[i];
                itemsDict[$"{itemName}[{i}]"] = currentItem;

                // 객체 속성 접근을 위한 인덱스 기반 참조 설정
                if (currentItem != null)
                {
                    var properties = currentItem.GetType().GetProperties();
                    foreach (var prop in properties)
                    {
                        if (prop.CanRead)
                        {
                            try
                            {
                                var value = prop.GetValue(currentItem);
                                itemsDict[$"{itemName}[{i}].{prop.Name}"] = value;
                            }
                            catch (Exception ex)
                            {
                                Logger.Warning($"Error getting property {prop.Name}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            else
            {
                // For indexes beyond available items, provide empty values to avoid errors
                itemsDict[$"{itemName}[{i}]"] = null;
            }
        }

        // Add the items dictionary to the context variables
        foreach (var kvp in itemsDict)
        {
            _context.Variables[kvp.Key] = kvp.Value;
        }

        // Additionally, set the first item as the named item for compatibility
        if (slideItems.Count > 0)
        {
            _context.Variables[itemName] = slideItems[0];

            // 슬라이드 내에서 전체 컬렉션 참조가 가능하도록 설정
            _context.Variables[$"{itemName}s"] = slideItems;

            // 추가 개선: 자주 사용되는 컬렉션 접근 방식
            _context.Variables["batch_items"] = slideItems;
            _context.Variables["current_item"] = slideItems[0];
        }

        // Process text replacements for this slide
        ProcessTextReplacements(slidePart);
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