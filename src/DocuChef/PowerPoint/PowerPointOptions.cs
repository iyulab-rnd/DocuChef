namespace DocuChef.PowerPoint;

/// <summary>
/// Configuration options specific to PowerPoint templates
/// </summary>
public class PowerPointOptions
{
    /// <summary>
    /// Plugins for PowerPoint presentation processing
    /// </summary>
    public List<IPowerPointPlugin> Plugins { get; set; } = [];

    /// <summary>
    /// Whether to update slide numbers
    /// </summary>
    public bool UpdateSlideNumbers { get; set; } = true;

    /// <summary>
    /// Whether to include hidden slides in the output
    /// </summary>
    public bool IncludeHiddenSlides { get; set; } = false;

    /// <summary>
    /// Whether to support dollar sign syntax (${variable})
    /// </summary>
    public bool SupportDollarSignSyntax { get; set; } = true;

    /// <summary>
    /// Additional namespaces to include in expression evaluation
    /// </summary>
    public List<string> AdditionalNamespaces { get; set; } = [];
}