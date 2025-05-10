namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Handles shapes with out-of-range array references
/// </summary>
internal class OutOfRangeShapeHandler
{
    private readonly PowerPointContext _context;
    private readonly HashSet<string> _hiddenShapeIds = new();

    /// <summary>
    /// Initialize handler
    /// </summary>
    public OutOfRangeShapeHandler(PowerPointContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Scan and hide all shapes with out-of-range array references
    /// </summary>
    public void ScanAndHideOutOfRangeShapes(SlidePart slidePart, Dictionary<string, int> arrayLengths)
    {
        if (slidePart?.Slide == null || !arrayLengths.Any())
            return;

        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        Logger.Debug($"Scanning {shapes.Count} shapes for out-of-range array references");

        // 범위를 벗어나는 도형 컬렉션
        var outOfRangeShapes = new List<(P.Shape Shape, string ShapeName, string ArrayName, int Index)>();

        // 모든 도형 검사
        foreach (var shape in shapes)
        {
            string shapeName = shape.GetShapeName();
            string shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString() ?? "(no id)";

            // 이미 숨겨진 도형은 건너뛰기
            if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value == true)
            {
                Logger.Debug($"Skipping already hidden shape: '{shapeName ?? "(unnamed)"}' (ID: {shapeId})");
                continue;
            }

            // 텍스트가 없는 도형은 건너뛰기
            if (shape.TextBody == null)
                continue;

            string text = GetShapeText(shape);
            if (string.IsNullOrEmpty(text))
                continue;

            // 각 배열에 대해 범위 검사
            foreach (var arrayEntry in arrayLengths)
            {
                string arrayName = arrayEntry.Key;
                int arrayLength = arrayEntry.Value;

                // 이 배열에 대한 참조가 있는지 확인
                if (!text.Contains($"{arrayName}["))
                    continue;

                Logger.Debug($"Checking shape '{shapeName ?? "(unnamed)"}' for {arrayName} references");

                // 모든 배열 인덱스 참조 찾기
                var matches = Regex.Matches(text, $"{arrayName}\\[(\\d+)\\]");
                foreach (Match match in matches)
                {
                    if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
                    {
                        // 인덱스가 배열 길이를 벗어나는지 확인
                        if (index >= arrayLength)
                        {
                            Logger.Debug($"Shape '{shapeName ?? "(unnamed)"}' references out-of-range index {index} for array {arrayName} (length: {arrayLength})");
                            outOfRangeShapes.Add((shape, shapeName, arrayName, index));
                            break; // 이 도형에 대해 더 이상 확인할 필요 없음
                        }
                    }
                }
            }
        }

        // 범위를 벗어나는 도형 숨기기
        if (outOfRangeShapes.Any())
        {
            Logger.Info($"Found {outOfRangeShapes.Count} shapes with out-of-range array references");

            foreach (var (shape, shapeName, arrayName, index) in outOfRangeShapes)
            {
                string shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString() ?? "(no id)";
                Logger.Debug($"Hiding shape '{shapeName ?? "(unnamed)"}' (ID: {shapeId}) with out-of-range index {index} for array {arrayName}");

                HideShapeCompletely(shape);
                _hiddenShapeIds.Add(shapeId);
            }

            // 슬라이드를 저장하여 변경사항 적용
            try
            {
                slidePart.Slide.Save();
                Logger.Debug("Saved slide after hiding out-of-range shapes");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error saving slide after hiding shapes: {ex.Message}", ex);
            }
        }
        else
        {
            Logger.Debug("No shapes with out-of-range array references found");
        }
    }

    /// <summary>
    /// 도형에서 텍스트 추출
    /// </summary>
    private string GetShapeText(P.Shape shape)
    {
        if (shape?.TextBody == null)
            return string.Empty;

        var sb = new StringBuilder();

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement != null && !string.IsNullOrEmpty(textElement.Text))
                {
                    sb.Append(textElement.Text);
                }
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// 도형을 완전히 숨기는 메서드
    /// </summary>
    private void HideShapeCompletely(P.Shape shape)
    {
        if (shape == null)
            return;

        try
        {
            string shapeName = shape.GetShapeName();
            string shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString() ?? "(no id)";

            // 1. Hidden 속성 설정
            var nvProps = shape.NonVisualShapeProperties;
            if (nvProps?.NonVisualDrawingProperties != null)
            {
                var nvDrawProps = nvProps.NonVisualDrawingProperties;
                nvDrawProps.Hidden = new BooleanValue(true);
            }

            // 2. 크기를 1로 설정
            if (shape.ShapeProperties?.Transform2D?.Extents != null)
            {
                // 원래 크기 백업
                var extents = shape.ShapeProperties.Transform2D.Extents;
                var nvAppProps = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;

                if (nvAppProps != null)
                {
                    long cx = extents.Cx?.Value ?? 0;
                    long cy = extents.Cy?.Value ?? 0;

                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originalcx", "", cx.ToString()));
                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originalcy", "", cy.ToString()));

                    // 크기를 1로 설정
                    extents.Cx = 1;
                    extents.Cy = 1;
                }
            }

            // 3. 위치를 슬라이드 바깥으로 이동
            if (shape.ShapeProperties?.Transform2D?.Offset != null)
            {
                var offset = shape.ShapeProperties.Transform2D.Offset;
                var nvAppProps = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;

                if (nvAppProps != null)
                {
                    long x = offset.X?.Value ?? 0;
                    long y = offset.Y?.Value ?? 0;

                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originalx", "", x.ToString()));
                    nvAppProps.SetAttribute(new OpenXmlAttribute("", "originaly", "", y.ToString()));

                    // 위치를 멀리 이동
                    offset.X = -10000000;
                    offset.Y = -10000000;
                }
            }

            // 4. 텍스트 지우기
            if (shape.TextBody != null)
            {
                shape.ClearText();
            }

            Logger.Debug($"Successfully hidden shape '{shapeName ?? "(unnamed)"}' (ID: {shapeId})");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error hiding shape: {ex.Message}");
        }
    }

    /// <summary>
    /// 숨겨진 도형 수 반환
    /// </summary>
    public int GetHiddenShapeCount()
    {
        return _hiddenShapeIds.Count;
    }

    /// <summary>
    /// 도형이 숨겨졌는지 확인
    /// </summary>
    public bool IsShapeHidden(string shapeId)
    {
        return !string.IsNullOrEmpty(shapeId) && _hiddenShapeIds.Contains(shapeId);
    }
}