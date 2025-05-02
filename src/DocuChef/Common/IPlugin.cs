namespace DocuChef.Common;

/// <summary>
/// Base interface for all document processing plugins
/// </summary>
public interface IPlugin
{
    /// <summary>
    /// Gets the name of the plugin
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Gets the description of the plugin
    /// </summary>
    string Description { get; }
}