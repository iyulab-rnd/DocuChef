using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocuChef.Utils;

namespace DocuChef.Word;

public partial class WordRecipe
{
    private async Task ProcessSectionsAsync()
    {
        // Get main document content
        var mainPart = _wordDoc?.MainDocumentPart;
        if (mainPart == null) return;

        var body = mainPart.Document.Body;
        if (body == null) return;

        // Find all paragraphs
        var paragraphs = body.Descendants<Paragraph>().ToList();

        // Find section markers
        List<SectionInfo> sections = FindSections(paragraphs);

        // Process sections
        foreach (var section in sections)
        {
            if (section.Type.Equals("TableRows", StringComparison.OrdinalIgnoreCase))
            {
                await ProcessTableRowsSectionAsync(section);
            }
            else if (section.Type.Equals("ListItems", StringComparison.OrdinalIgnoreCase))
            {
                await ProcessListItemsSectionAsync(section);
            }
        }
    }

    private static List<SectionInfo> FindSections(List<Paragraph> paragraphs)
    {
        var sections = new List<SectionInfo>();
        SectionInfo? currentSection = null;

        foreach (var paragraph in paragraphs)
        {
            string text = paragraph.InnerText;

            if (text.Contains("<!--#begin:"))
            {
                // Parse begin marker
                var match = SectionRegex.Match(text);
                if (match.Success)
                {
                    string type = match.Groups[1].Value;
                    string name = match.Groups[2].Value;

                    currentSection = new SectionInfo(type, name, paragraph);
                }
            }
            else if (text.Contains("<!--#end:") && currentSection != null)
            {
                currentSection.EndParagraph = paragraph;
                sections.Add(currentSection);
                currentSection = null;
            }
            else if (currentSection != null)
            {
                currentSection.ContentParagraphs.Add(paragraph);
            }
        }

        return sections;
    }

    private async Task ProcessTableRowsSectionAsync(SectionInfo section)
    {
        try
        {
            // Get collection name
            string collectionName = section.Name;

            // Get collection data
            var collection = TextProcessingHelper.ResolveCollection(collectionName, Data, Options.VariableResolver);
            if (!collection.Any()) return;

            // Find the table containing the section
            var table = FindContainingTable(section.StartParagraph);
            if (table == null) return;

            // Find template row
            var templateRow = FindTemplateRow(table, section.StartParagraph);
            if (templateRow == null) return;

            // Clone and bind data for each collection item
            int insertIndex = GetRowIndex(table, templateRow) + 1;

            foreach (var item in collection)
            {
                // Create new row
                TableRow newRow = (TableRow)templateRow.CloneNode(true);

                // Process text elements in the row
                foreach (var text in newRow.Descendants<Text>())
                {
                    if (text.Text.Contains("${"))
                    {
                        // Create combined context with item properties
                        var combinedData = TextProcessingHelper.CreateCombinedContext(item, Data);
                        await ProcessTextElementWithDataAsync(text, combinedData);
                    }
                }

                // Insert row
                table.InsertAt(newRow, insertIndex++);
            }

            // Remove template row
            templateRow.Remove();

            // Remove marker paragraphs
            section.StartParagraph.Remove();
            section.EndParagraph.Remove();
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError($"Error processing table rows section '{section.Name}'", ex);
            throw new TemplateException($"Error processing table rows section: {ex.Message}", ex);
        }
    }

    private async Task ProcessListItemsSectionAsync(SectionInfo section)
    {
        // This is a placeholder for list items processing
        // In the original code, this was not implemented
        LoggingHelper.LogWarning($"List items section processing not implemented for section '{section.Name}'");
        await Task.CompletedTask; // Ensure async method
    }

    private static Table? FindContainingTable(Paragraph paragraph)
    {
        // Find parent table
        OpenXmlElement? parent = paragraph.Parent;
        while (parent != null && parent is not Table)
        {
            parent = parent.Parent;
        }

        return parent as Table;
    }

    private static TableRow? FindTemplateRow(Table table, Paragraph markerParagraph)
    {
        // Find row containing the marker paragraph
        foreach (var row in table.Elements<TableRow>())
        {
            foreach (var cell in row.Elements<TableCell>())
            {
                if (cell.Descendants<Paragraph>().Contains(markerParagraph))
                {
                    return row;
                }
            }
        }

        return null;
    }

    private static int GetRowIndex(Table table, TableRow row)
    {
        int index = 0;
        foreach (var r in table.Elements<TableRow>())
        {
            if (r == row)
                return index;
            index++;
        }
        return -1;
    }

    private class SectionInfo
    {
        public string Type { get; }
        public string Name { get; }
        public Paragraph StartParagraph { get; }
        public Paragraph EndParagraph { get; set; }
        public List<Paragraph> ContentParagraphs { get; }

        public SectionInfo(string type, string name, Paragraph startParagraph)
        {
            Type = type ?? throw new ArgumentNullException(nameof(type));
            Name = name ?? throw new ArgumentNullException(nameof(name));
            StartParagraph = startParagraph ?? throw new ArgumentNullException(nameof(startParagraph));
            EndParagraph = startParagraph; // Will be updated later
            ContentParagraphs = new List<Paragraph>();
        }
    }
}