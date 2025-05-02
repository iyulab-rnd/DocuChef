using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocuChef.PowerPoint;

public partial class PowerPointRecipe
{
    private async Task ProcessSlideContentAsync(SlidePart slidePart)
    {
        // Process all text elements in the slide
        foreach (var textElement in slidePart.Slide.Descendants<A.Text>())
        {
            await ProcessTextElementAsync(textElement);
        }
    }

    private async Task ProcessSlideContentWithContextAsync(SlidePart slidePart, object context)
    {
        // Combine global data with context
        var combinedData = TextProcessingHelper.CreateCombinedContext(context, Data);

        // Process text elements
        foreach (var textElement in slidePart.Slide.Descendants<A.Text>())
        {
            await ProcessTextElementWithDataAsync(textElement, combinedData);
        }
    }

    private async Task ProcessTextElementAsync(A.Text text)
    {
        await ProcessTextElementWithDataAsync(text, Data);
    }

    private async Task ProcessTextElementWithDataAsync(A.Text text, Dictionary<string, object> contextData)
    {
        try
        {
            // Skip empty text
            if (string.IsNullOrEmpty(text.Text)) return;

            // Skip if no variables
            if (!text.Text.Contains("${") && !text.Text.Contains('{')) return;

            // Use common text processing helper
            string processed = await TextProcessingHelper.ProcessVariablesAsync(
                text.Text,
                contextData,
                Options.CultureInfo,
                Options.PowerPoint.SupportDollarSignSyntax,
                (expr, obj) => Options.VariableResolver?.Invoke(expr, contextData),
                Options.PowerPoint.AdditionalNamespaces);

            text.Text = processed;
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError($"Error processing text element '{text.Text}'", ex);
            throw new TemplateException($"Error processing text element '{text.Text}': {ex.Message}", ex);
        }
    }

    private string GetNotesText(NotesSlidePart notesPart)
    {
        if (notesPart?.NotesSlide == null) return string.Empty;

        var textBuilder = new StringBuilder();

        foreach (var textElement in notesPart.NotesSlide.Descendants<A.Text>())
        {
            if (textElement != null)
            {
                textBuilder.Append(textElement.Text);
            }
        }

        return textBuilder.ToString();
    }
}

/// <summary>
/// Class representing a slide directive from slide notes
/// </summary>
internal class SlideDirective
{
    /// <summary>
    /// Gets the directive type (e.g., 'repeat', 'if')
    /// </summary>
    public string Type { get; }

    /// <summary>
    /// Gets the directive value
    /// </summary>
    public string Value { get; }

    /// <summary>
    /// Creates a new slide directive
    /// </summary>
    public SlideDirective(string type, string value)
    {
        Type = type ?? throw new ArgumentNullException(nameof(type));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }
}