namespace DocuChef.Word;

/// <summary>
/// Configuration options specific to Word templates
/// </summary>
public class WordOptions
{
    /// <summary>
    /// Plugins for Word document processing
    /// </summary>
    public List<IWordPlugin> Plugins { get; set; } = [];

    /// <summary>
    /// Whether to update fields after data binding
    /// </summary>
    public bool UpdateFieldsAfterBinding { get; set; } = true;

    /// <summary>
    /// Whether to update table of contents
    /// </summary>
    public bool UpdateTableOfContents { get; set; } = false;

    /// <summary>
    /// Whether to process content in headers and footers
    /// </summary>
    public bool ProcessHeadersAndFooters { get; set; } = true;

    /// <summary>
    /// Whether to preserve formatting during text replacement
    /// </summary>
    public bool PreserveFormatting { get; set; } = true;

    public List<string> AdditionalNamespaces { get; set; } = [];
}