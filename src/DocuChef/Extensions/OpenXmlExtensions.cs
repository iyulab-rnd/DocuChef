namespace DocuChef.Extensions;

/// <summary>
/// Extension methods for working with OpenXml objects
/// </summary>
public static class OpenXmlExtensions
{
    /// <summary>
    /// Gets the name of a PowerPoint shape, prioritizing Alt Text over default shape name
    /// </summary>
    public static string GetShapeName(this Shape shape)
    {
        if (shape?.NonVisualShapeProperties == null)
        {
            Logger.Debug("Shape or NonVisualShapeProperties is null.");
            return null;
        }

        string shapeName = null;

        // NonVisualDrawingProperties에서 속성 확인
        var nvdp = shape.NonVisualShapeProperties.NonVisualDrawingProperties;
        if (nvdp != null)
        {
            if (!string.IsNullOrWhiteSpace(nvdp.Name?.Value))
            {
                Logger.Debug($"Found shape name attribute: {nvdp.Name.Value}");
                shapeName = nvdp.Name.Value;
            }
        }
        else
        {
            Logger.Debug("NonVisualDrawingProperties is null.");
        }

        if (shapeName == null)
        {
            Logger.Warning("No valid shape name or alt text found.");
        }

        return shapeName;
    }

    /// <summary>
    /// Gets the text content of a PowerPoint shape
    /// </summary>
    public static string GetText(this P.Shape shape)
    {
        if (shape == null)
            return string.Empty;

        var paragraphs = shape.Descendants<A.Paragraph>();
        if (!paragraphs.Any())
            return string.Empty;

        var sb = new StringBuilder();

        foreach (var paragraph in paragraphs)
        {
            if (sb.Length > 0)
                sb.AppendLine();

            foreach (var run in paragraph.Elements<A.Run>())
            {
                var text = run.GetFirstChild<A.Text>();
                if (text != null)
                {
                    sb.Append(text.Text);
                }
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Gets the text content of a PowerPoint shape with formatting information
    /// </summary>
    public static List<(string Text, A.RunProperties Properties)> GetFormattedText(this P.Shape shape)
    {
        var result = new List<(string Text, A.RunProperties Properties)>();

        if (shape?.TextBody == null)
            return result;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var text = run.GetFirstChild<A.Text>();
                if (text != null && !string.IsNullOrEmpty(text.Text))
                {
                    var props = run.RunProperties?.CloneNode(true) as A.RunProperties;
                    result.Add((text.Text, props));
                }
            }

            // Add paragraph break marker if not the last paragraph
            if (paragraph != shape.TextBody.Elements<A.Paragraph>().Last())
            {
                result.Add(("\n", null));
            }
        }

        return result;
    }

    /// <summary>
    /// Clears text content from a PowerPoint shape
    /// </summary>
    public static void ClearText(this P.Shape shape)
    {
        if (shape?.TextBody == null)
            return;

        // Keep only one paragraph with one empty run
        var textBody = shape.TextBody;

        // Remove all paragraphs
        var paragraphs = textBody.Elements<A.Paragraph>().ToList();
        foreach (var para in paragraphs)
        {
            para.Remove();
        }

        // Add a single empty paragraph
        var emptyParagraph = new A.Paragraph();
        var emptyRun = new A.Run();
        emptyRun.AppendChild(new A.Text());
        emptyParagraph.AppendChild(emptyRun);
        textBody.AppendChild(emptyParagraph);
    }

    /// <summary>
    /// Sets text content in a PowerPoint shape, preserving formatting by processing each run individually
    /// </summary>
    public static bool SetTextWithExpressions(this P.Shape shape, IExpressionEvaluator processor, Dictionary<string, object> variables)
    {
        if (shape?.TextBody == null || processor == null)
            return false;

        bool hasChanges = false;

        // Process each paragraph
        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            // Process each run
            foreach (var run in paragraph.Elements<A.Run>().ToList())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string originalText = textElement.Text;

                // Check if there are expressions to process
                if (!originalText.Contains("${"))
                    continue;

                // Process expressions
                string processedText = ProcessExpressions(originalText, processor, variables);

                // Skip if no changes
                if (processedText == originalText)
                    continue;

                // Update text
                textElement.Text = processedText;
                hasChanges = true;
            }
        }

        return hasChanges;
    }

    /// <summary>
    /// Process expressions in text
    /// </summary>
    private static string ProcessExpressions(string text, IExpressionEvaluator processor, Dictionary<string, object> variables)
    {
        if (string.IsNullOrEmpty(text) || processor == null)
            return text;

        try
        {
            // Process ${...} expressions using regex
            var regex = new Regex(@"\${([^{}]+)}");
            return regex.Replace(text, match => {
                try
                {
                    // Evaluate the expression
                    var expressionValue = processor.EvaluateCompleteExpression(match.Value, variables);
                    return expressionValue?.ToString() ?? "";
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Error evaluating expression '{match.Value}': {ex.Message}");
                    return match.Value; // Keep original on error
                }
            });
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing expressions: {ex.Message}", ex);
            return text;
        }
    }

    /// <summary>
    /// Sets text content in a PowerPoint shape, preserving formatting
    /// </summary>
    public static void SetText(this P.Shape shape, string text)
    {
        if (shape?.TextBody == null)
            return;

        try
        {
            Logger.Debug($"Setting text in shape: {text}");

            // Clear existing text but preserve paragraph properties
            var textBody = shape.TextBody;
            var existingParagraphs = textBody.Elements<A.Paragraph>().ToList();

            // Backup paragraph properties and formatting
            var paragraphProperties = existingParagraphs
                .Select(p => p.ParagraphProperties?.CloneNode(true) as A.ParagraphProperties)
                .ToList();

            // Backup first run properties for formatting preservation
            A.RunProperties runProps = null;
            foreach (var para in existingParagraphs)
            {
                var firstRun = para.Elements<A.Run>().FirstOrDefault();
                if (firstRun?.RunProperties != null)
                {
                    runProps = firstRun.RunProperties.CloneNode(true) as A.RunProperties;
                    break;
                }
            }

            // Remove all existing paragraphs
            foreach (var para in existingParagraphs)
            {
                para.Remove();
            }

            // Split text into lines
            string[] lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.None);
            if (lines.Length == 0)
            {
                lines = new[] { string.Empty };
            }

            // Create new paragraphs
            for (int i = 0; i < lines.Length; i++)
            {
                var para = new A.Paragraph();

                // Apply preserved paragraph properties if available
                if (i < paragraphProperties.Count && paragraphProperties[i] != null)
                {
                    para.AppendChild(paragraphProperties[i]);
                }

                var run = new A.Run();

                // Apply preserved run properties if available
                if (runProps != null)
                {
                    run.RunProperties = runProps.CloneNode(true) as A.RunProperties;
                }

                run.AppendChild(new A.Text(lines[i]));
                para.AppendChild(run);
                textBody.AppendChild(para);
            }

            Logger.Debug($"Text set successfully in shape");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error setting text in shape: {ex.Message}", ex);

            // Fallback method for setting text
            try
            {
                var textBody = shape.TextBody;

                // Clear all existing text elements
                var textElements = textBody.Descendants<A.Text>().ToList();
                foreach (var textElement1 in textElements)
                {
                    textElement1.Text = "";
                }

                // Get or create a paragraph and run
                var paragraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
                if (paragraph == null)
                {
                    paragraph = new A.Paragraph();
                    textBody.AppendChild(paragraph);
                }

                var run = paragraph.Elements<A.Run>().FirstOrDefault();
                if (run == null)
                {
                    run = new A.Run();
                    paragraph.AppendChild(run);
                }

                var textElement2 = run.Elements<A.Text>().FirstOrDefault();
                if (textElement2 == null)
                {
                    textElement2 = new A.Text();
                    run.AppendChild(textElement2);
                }

                // Set the text
                textElement2.Text = text;
                Logger.Debug($"Text set with fallback method in shape");
            }
            catch (Exception fallbackEx)
            {
                Logger.Error($"Fallback method also failed: {fallbackEx.Message}", fallbackEx);
            }
        }
    }

    /// <summary>
    /// Sets visibility of a PowerPoint shape
    /// </summary>
    public static void SetVisibility(this P.Shape shape, bool visible)
    {
        if (shape == null)
            return;

        // Ensure ShapeProperties exists
        if (shape.ShapeProperties == null)
            shape.ShapeProperties = new P.ShapeProperties();

        // Ensure Transform2D exists
        if (shape.ShapeProperties.Transform2D == null)
            shape.ShapeProperties.Transform2D = new A.Transform2D();

        // Get or create NonVisualProperties
        var nvProps = shape.NonVisualShapeProperties;
        if (nvProps == null)
            return;

        if (!visible)
        {
            // Store original dimensions if they exist
            var transform = shape.ShapeProperties.Transform2D;
            if (transform?.Extents != null)
            {
                long cx = transform.Extents.Cx?.Value ?? 0;
                long cy = transform.Extents.Cy?.Value ?? 0;

                // Store original dimensions as custom attributes for later restoration
                nvProps.ApplicationNonVisualDrawingProperties.SetAttribute(
                    new OpenXmlAttribute("", "originalcx", "", cx.ToString()));
                nvProps.ApplicationNonVisualDrawingProperties.SetAttribute(
                    new OpenXmlAttribute("", "originalcy", "", cy.ToString()));

                // Set zero size to hide
                transform.Extents.Cx = 0;
                transform.Extents.Cy = 0;
            }
        }
        else
        {
            // Restore original dimensions if previously hidden
            var nvAppProps = nvProps.ApplicationNonVisualDrawingProperties;
            var origCxAttr = nvAppProps.GetAttributes().FirstOrDefault(a => a.LocalName == "originalcx");
            var origCyAttr = nvAppProps.GetAttributes().FirstOrDefault(a => a.LocalName == "originalcy");

            if (origCxAttr.Value != null && origCyAttr.Value != null)
            {
                var transform = shape.ShapeProperties.Transform2D;
                if (transform?.Extents == null)
                    transform.Extents = new A.Extents();

                if (long.TryParse(origCxAttr.Value, out long cx) &&
                    long.TryParse(origCyAttr.Value, out long cy))
                {
                    transform.Extents.Cx = cx;
                    transform.Extents.Cy = cy;
                }
            }
        }
    }

    /// <summary>
    /// Gets notes from a PowerPoint slide
    /// </summary>
    public static string GetNotes(this SlidePart slidePart)
    {
        if (slidePart?.NotesSlidePart?.NotesSlide == null)
        {
            Logger.Debug("NotesSlidePart or NotesSlide is null");
            return string.Empty;
        }

        Logger.Debug("Found NotesSlide");

        // 모든 Shape 요소 찾기
        var shapes = slidePart.NotesSlidePart.NotesSlide.Descendants<P.Shape>().ToList();
        Logger.Debug($"Found {shapes.Count} shapes in notes slide");

        // 모든 텍스트 요소 로깅
        foreach (var shape in shapes)
        {
            var texts = shape.Descendants<A.Text>().Select(t => t.Text).Where(t => !string.IsNullOrEmpty(t)).ToList();
            Logger.Debug($"Shape ID: {shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}, Texts: {string.Join(", ", texts)}");
        }

        // 모든 텍스트 수집
        var allTexts = slidePart.NotesSlidePart.NotesSlide
            .Descendants<A.Text>()
            .Select(t => t.Text)
            .Where(t => !string.IsNullOrEmpty(t))
            .ToList();

        Logger.Debug($"All texts in notes slide: {string.Join(", ", allTexts)}");

        // 직접적인 방법: 지시문 형식의 텍스트 찾기
        var directiveText = allTexts.FirstOrDefault(t => t.StartsWith("#"));
        if (!string.IsNullOrEmpty(directiveText))
        {
            Logger.Debug($"Found directive text: {directiveText}");
            return directiveText;
        }

        // 그 외의 경우 모든 텍스트 결합 (단, 숫자만 있는 텍스트는 제외)
        var result = string.Join(" ", allTexts.Where(t => !System.Text.RegularExpressions.Regex.IsMatch(t, @"^\d+$")));
        Logger.Debug($"Returning combined text: {result}");
        return result;
    }

    /// <summary>
    /// Creates a copy of an OpenXml part
    /// </summary>
    public static void CopyTo(this OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        using (var stream = sourcePart.GetStream())
        {
            stream.Position = 0;
            targetPart.FeedData(stream);
        }
    }
}