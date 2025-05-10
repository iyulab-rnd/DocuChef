namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Manages visibility of shapes based on data availability
/// </summary>
internal class ShapeVisibilityManager
{
    private readonly PowerPointContext _context;

    /// <summary>
    /// Initialize shape visibility manager
    /// </summary>
    public ShapeVisibilityManager(PowerPointContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Hide shapes that reference array indices beyond available data range
    /// </summary>
    public void HideShapesWithOutOfRangeIndices(SlidePart slidePart, string arrayName, int availableItemCount, int startIndex, int itemsPerSlide)
    {
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        Logger.Debug($"Checking visibility for {shapes.Count} shapes based on {arrayName} data availability");

        // 도형 ID와 텍스트 내용을 기록하는 딕셔너리 (로깅 목적)
        var shapeTexts = new Dictionary<string, string>();

        foreach (var shape in shapes)
        {
            string shapeName = shape.GetShapeName();
            string shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString() ?? "(no id)";

            // 텍스트 내용 수집
            if (shape.TextBody != null)
            {
                var textBuilder = new StringBuilder();
                foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
                {
                    foreach (var run in paragraph.Elements<A.Run>())
                    {
                        var textElement = run.GetFirstChild<A.Text>();
                        if (textElement != null && !string.IsNullOrEmpty(textElement.Text))
                        {
                            textBuilder.Append(textElement.Text);
                        }
                    }
                }
                shapeTexts[shapeId] = textBuilder.ToString();
            }

            // Skip shapes without text body
            if (shape.TextBody == null)
                continue;

            // Check if shape contains references to out-of-range indices
            if (ContainsOutOfRangeArrayReferences(shape, arrayName, availableItemCount, startIndex, itemsPerSlide))
            {
                Logger.Debug($"Hiding shape '{shapeName ?? "(unnamed)"}' (ID: {shapeId}) with out-of-range array references");
                Logger.Debug($"Shape text content: {shapeTexts.GetValueOrDefault(shapeId, "(empty)")}");
                SetShapeVisibility(shape, false);
            }
            else
            {
                // Check if this shape contains any array references at all
                if (ContainsArrayReference(shape, arrayName))
                {
                    Logger.Debug($"Ensuring shape '{shapeName ?? "(unnamed)"}' (ID: {shapeId}) is visible as it contains in-range references");
                    Logger.Debug($"Shape text content: {shapeTexts.GetValueOrDefault(shapeId, "(empty)")}");
                    SetShapeVisibility(shape, true);
                }
            }
        }
    }

    /// <summary>
    /// Check if shape contains array reference to the specified array
    /// </summary>
    private bool ContainsArrayReference(P.Shape shape, string arrayName)
    {
        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;

                if (text.Contains($"${{{arrayName}[") || text.Contains($"{arrayName}["))
                    return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Check if shape contains references to array indices beyond available data
    /// </summary>
    private bool ContainsOutOfRangeArrayReferences(P.Shape shape, string arrayName, int availableItemCount, int startIndex, int itemsPerSlide)
    {
        // 텍스트에서 ${arrayName[index]} 형태의 표현식을 찾기 위한 패턴
        var pattern = $"\\${{{arrayName}\\[(\\d+)\\]";

        // 이미지 함수에서의 배열 참조도 확인하는 추가 패턴
        var imageFunctionPattern = $"\\${{ppt\\.Image\\({arrayName}\\[(\\d+)\\]";

        // 또 다른 패턴: ppt.Image(arrayName[index])
        var simpleFunctionPattern = $"ppt\\.Image\\({arrayName}\\[(\\d+)\\]";

        // 모든 텍스트 실행에서 확인
        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;

                // 일반 배열 참조 표현식 찾기
                var matches = Regex.Matches(text, pattern);
                foreach (Match match in matches)
                {
                    if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
                    {
                        // 중요: 인덱스가 사용 가능한 데이터 수를 초과하는지 확인
                        if (index >= availableItemCount)
                        {
                            Logger.Debug($"Shape contains out-of-range index: {index} (available: {availableItemCount})");
                            return true;
                        }
                    }
                }

                // 이미지 함수에서의 배열 참조 확인
                var functionMatches = Regex.Matches(text, imageFunctionPattern);
                foreach (Match match in functionMatches)
                {
                    if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
                    {
                        // 이미지 함수에서의 인덱스가 범위를 벗어나는지 확인
                        if (index >= availableItemCount)
                        {
                            Logger.Debug($"Shape contains image function with out-of-range index: {index} (available: {availableItemCount})");
                            return true;
                        }
                    }
                }

                // 단순 이미지 함수 패턴 확인
                var simpleFunctionMatches = Regex.Matches(text, simpleFunctionPattern);
                foreach (Match match in simpleFunctionMatches)
                {
                    if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
                    {
                        if (index >= availableItemCount)
                        {
                            Logger.Debug($"Shape contains simple image function with out-of-range index: {index} (available: {availableItemCount})");
                            return true;
                        }
                    }
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Set shape visibility 
    /// </summary>
    private void SetShapeVisibility(P.Shape shape, bool visible)
    {
        if (shape == null)
            return;

        try
        {
            // Method 1: Set hidden attribute (most reliable)
            var nvProps = shape.NonVisualShapeProperties;
            if (nvProps != null)
            {
                var nvDrawProps = nvProps.NonVisualDrawingProperties;
                if (nvDrawProps != null)
                {
                    nvDrawProps.Hidden = !visible;
                    Logger.Debug($"Set Hidden={!visible} on shape");
                }
            }

            // Method 2: Make the shape very small (backup method)
            if (!visible && shape.ShapeProperties?.Transform2D?.Extents != null)
            {
                // Store original dimensions for potential future use
                var extents = shape.ShapeProperties.Transform2D.Extents;
                var nvAppProps = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;

                if (nvAppProps != null)
                {
                    long cx = extents.Cx?.Value ?? 0;
                    long cy = extents.Cy?.Value ?? 0;

                    // Only store if not already stored
                    var cxAttr = nvAppProps.GetAttributes().FirstOrDefault(a => a.LocalName == "originalcx");
                    var cyAttr = nvAppProps.GetAttributes().FirstOrDefault(a => a.LocalName == "originalcy");

                    if (cxAttr.Value == null && cyAttr.Value == null && cx > 0 && cy > 0)
                    {
                        // Store original dimensions as custom attributes
                        nvAppProps.SetAttribute(
                            new OpenXmlAttribute("", "originalcx", "", cx.ToString()));
                        nvAppProps.SetAttribute(
                            new OpenXmlAttribute("", "originalcy", "", cy.ToString()));

                        Logger.Debug($"Stored original dimensions: {cx}x{cy}");
                    }

                    // Set to almost zero size
                    extents.Cx = 1;
                    extents.Cy = 1;
                    Logger.Debug("Set shape to minimum size");
                }
            }
            else if (visible)
            {
                // Restore original dimensions if previously hidden
                var nvAppProps = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;
                if (nvAppProps != null)
                {
                    var origCxAttr = nvAppProps.GetAttributes().FirstOrDefault(a => a.LocalName == "originalcx");
                    var origCyAttr = nvAppProps.GetAttributes().FirstOrDefault(a => a.LocalName == "originalcy");

                    if (origCxAttr.Value != null && origCyAttr.Value != null)
                    {
                        var transform = shape.ShapeProperties?.Transform2D;
                        if (transform?.Extents != null &&
                            long.TryParse(origCxAttr.Value, out long cx) &&
                            long.TryParse(origCyAttr.Value, out long cy))
                        {
                            transform.Extents.Cx = cx;
                            transform.Extents.Cy = cy;
                            Logger.Debug($"Restored original dimensions: {cx}x{cy}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Could not set shape visibility: {ex.Message}");
        }
    }
}