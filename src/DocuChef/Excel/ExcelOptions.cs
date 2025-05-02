namespace DocuChef.Excel;

public class ExcelOptions
{
    /// <summary>
    /// Plugins for Excel document processing
    /// </summary>
    public List<IExcelPlugin> Plugins { get; set; } = new List<IExcelPlugin>();

    /// <summary>
    /// Whether to adjust column widths to content
    /// </summary>
    public bool AutoFitColumns { get; set; } = false;

    /// <summary>
    /// Whether to process cells with formulas
    /// </summary>
    public bool ProcessFormulaCells { get; set; } = true;
}