using DocuChef.Common;

namespace DocuChef.Excel;

/// <summary>
/// Interface for Excel document processing plugins
/// </summary>
public interface IExcelPlugin : IPlugin
{
    /// <summary>
    /// Executes the plugin on the specified workbook
    /// </summary>
    void Execute(object workbook, object data, RecipeOptions options);
}