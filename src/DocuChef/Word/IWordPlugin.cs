using DocuChef.Common;

namespace DocuChef.Word;

/// <summary>
/// Interface for Word document processing plugins
/// </summary>
public interface IWordPlugin : IPlugin
{
    /// <summary>
    /// Executes the plugin on the specified document
    /// </summary>
    void Execute(object document, object data, RecipeOptions options);
}