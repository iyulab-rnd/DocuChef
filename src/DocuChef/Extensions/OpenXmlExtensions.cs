using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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

        var textElements = shape.Descendants<A.Text>();
        return string.Join(" ", textElements.Select(t => t.Text ?? string.Empty));
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

        // Set text as not editable
        var bodyProps = textBody.GetFirstChild<A.BodyProperties>();
        if (bodyProps == null)
        {
            bodyProps = new A.BodyProperties();
            textBody.InsertAt(bodyProps, 0);
        }

        bodyProps.SetAttribute(new OpenXmlAttribute("noTextEdit", null, "1"));
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