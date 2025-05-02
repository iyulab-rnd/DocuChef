using DocuChef.Common;
using DocumentFormat.OpenXml.Packaging;

namespace DocuChef.PowerPoint;

/// <summary>
/// Interface for PowerPoint presentation processing plugins
/// </summary>
public interface IPowerPointPlugin : IPlugin
{
    /// <summary>
    /// Executes the plugin on the specified presentation
    /// </summary>
    void Execute(PresentationDocument presentation, Dictionary<string, object> data, RecipeOptions options);
}