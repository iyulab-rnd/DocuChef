using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office.Word;

namespace DocuChef.Word;

public partial class WordRecipe
{
    private async Task ProcessMainDocumentAsync()
    {
        if (Document == null)
        {
            throw new InvalidOperationException("Document is not initialized.");
        }

        // Get all paragraphs, runs, and text elements
        var paragraphs = Document.Descendants<Paragraph>().ToList();

        foreach (var paragraph in paragraphs)
        {
            await ProcessParagraphAsync(paragraph);
        }

        // Process tables
        var tables = Document.Descendants<Table>().ToList();
        foreach (var table in tables)
        {
            await ProcessTableAsync(table);
        }
    }

    private async Task ProcessParagraphAsync(Paragraph paragraph)
    {
        // Get the full text of the paragraph
        var fullText = paragraph.InnerText;

        // Check if the paragraph contains any variables
        if (!fullText.Contains("${") && !fullText.Contains('{'))
            return;

        // Process each run in the paragraph
        foreach (var run in paragraph.Elements<Run>().ToList())
        {
            foreach (var text in run.Elements<Text>().ToList())
            {
                if (text.Text.Contains("${") || text.Text.Contains('{'))
                {
                    await ProcessTextElementAsync(text);
                }
            }
        }
    }

    private async Task ProcessTextElementAsync(Text text)
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
                Data,
                Options.CultureInfo,
                true, // Support dollar sign syntax
                (expr, obj) => Options.VariableResolver?.Invoke(expr, Data),
                Options.Word.AdditionalNamespaces);

            text.Text = processed;
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError($"Error processing text element '{text.Text}'", ex);
            throw new TemplateException($"Error processing text element '{text.Text}': {ex.Message}", ex);
        }
    }

    private async Task ProcessTextElementWithDataAsync(Text text, Dictionary<string, object> contextData)
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
                Data,
                Options.CultureInfo,
                true, // Support dollar sign syntax
                (expr, obj) => Options.VariableResolver?.Invoke(expr, Data),
                Options.Word.AdditionalNamespaces);

            text.Text = processed;
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError($"Error processing text element '{text.Text}' with context data", ex);
            // Log but continue with other elements
        }
    }

    private async Task UpdateFieldsAsync()
    {
        if (_wordDoc?.MainDocumentPart == null) return;

        try
        {
            // Find all field code elements
            var fieldCodes = _wordDoc.MainDocumentPart.Document.Descendants<FieldCode>().ToList();

            foreach (var fieldCode in fieldCodes)
            {
                // Check field content
                var fieldText = fieldCode.Text?.Trim();
                if (string.IsNullOrEmpty(fieldText)) continue;

                // Handle special fields (TOC, DATE, etc.)
                if (fieldText.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                {
                    // TOC fields are handled in UpdateTableOfContents
                    continue;
                }

                if (fieldText.StartsWith("DATE", StringComparison.OrdinalIgnoreCase))
                {
                    // Update date fields to current date
                    await UpdateDateFieldAsync(fieldCode);
                }
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error updating fields", ex);
            throw new TemplateException($"Error updating fields: {ex.Message}", ex);
        }
    }

    private async Task UpdateDateFieldAsync(FieldCode fieldCode)
    {
        // Find the result node
        var resultNode = fieldCode.Parent?.Descendants<Run>()
            .FirstOrDefault(r => r.InnerText.Contains("DOCPROPERTY"));

        if (resultNode != null)
        {
            foreach (var textElement in resultNode.Elements<Text>())
            {
                textElement.Text = DateTime.Now.ToString(Options.CultureInfo);
            }
        }

        await Task.CompletedTask; // Ensure async method
    }

    private async Task UpdateTableOfContentsAsync()
    {
        if (_wordDoc?.MainDocumentPart == null) return;

        try
        {
            // Find all TOC fields
            var tocFields = _wordDoc.MainDocumentPart.Document.Descendants<FieldCode>()
                .Where(f => f.Text != null && f.Text.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (tocFields.Count == 0) return;

            // TOC fields are difficult to update directly
            // Set a flag to request automatic update when opening the document
            var settings = _wordDoc.MainDocumentPart.DocumentSettingsPart;
            if (settings != null)
            {
                var updateFields = new UpdateFields { Val = true };

                if (settings.Settings.Elements<UpdateFields>().Any())
                {
                    // Update existing element
                    settings.Settings.Elements<UpdateFields>().First().Val = true;
                }
                else
                {
                    // Add new element
                    settings.Settings.AppendChild(updateFields);
                }
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error updating table of contents", ex);
            throw new TemplateException($"Error updating table of contents: {ex.Message}", ex);
        }

        await Task.CompletedTask; // Ensure async method
    }

    private async Task ProcessTableAsync(Table table)
    {
        foreach (var row in table.Elements<TableRow>())
        {
            foreach (var cell in row.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    await ProcessParagraphAsync(paragraph);
                }
            }
        }
    }

    private async Task ProcessHeadersAndFootersAsync()
    {
        if (_wordDoc?.MainDocumentPart == null) return;

        var headerParts = _wordDoc.MainDocumentPart.HeaderParts;
        var footerParts = _wordDoc.MainDocumentPart.FooterParts;

        // Process headers
        foreach (var headerPart in headerParts)
        {
            var header = headerPart.Header;

            foreach (var paragraph in header.Descendants<Paragraph>())
            {
                await ProcessParagraphAsync(paragraph);
            }

            foreach (var table in header.Descendants<Table>())
            {
                await ProcessTableAsync(table);
            }
        }

        // Process footers
        foreach (var footerPart in footerParts)
        {
            var footer = footerPart.Footer;

            foreach (var paragraph in footer.Descendants<Paragraph>())
            {
                await ProcessParagraphAsync(paragraph);
            }

            foreach (var table in footer.Descendants<Table>())
            {
                await ProcessTableAsync(table);
            }
        }
    }

    private class UpdateFields : DocumentFormat.OpenXml.OpenXmlLeafElement
    {
        public bool? Val { get; set; }

        public UpdateFields() : base() { }

        public UpdateFields(bool val) : base()
        {
            Val = val;
        }
    }
}