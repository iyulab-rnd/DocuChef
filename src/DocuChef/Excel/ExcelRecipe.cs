using ClosedXML.Excel;
using ClosedXML.Report.XLCustom;
using DocuChef.Extensions;
using DocuChef.Logging;

namespace DocuChef.Excel;

/// <summary>
/// Represents an Excel template for document generation
/// </summary>
public class ExcelRecipe : RecipeBase
{
    private readonly ExcelOptions _options;
    private readonly XLCustomTemplate _template;
    private bool _isGenerated;

    /// <summary>
    /// Creates a new Excel template from a file
    /// </summary>
    public ExcelRecipe(string templatePath, ExcelOptions options = null)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        _options = options ?? new ExcelOptions();

        try
        {
            Logger.Debug($"Initializing Excel template from {templatePath}");
            _template = new XLCustomTemplate(templatePath, _options.TemplateOptions);

            InitializeTemplate();
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to initialize Excel template from {templatePath}", ex);
            throw new DocuChefException($"Failed to initialize Excel template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates a new Excel template from a stream
    /// </summary>
    public ExcelRecipe(Stream templateStream, ExcelOptions options = null)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        _options = options ?? new ExcelOptions();

        try
        {
            Logger.Debug("Initializing Excel template from stream");
            _template = new XLCustomTemplate(templateStream, _options.TemplateOptions);

            InitializeTemplate();
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to initialize Excel template from stream", ex);
            throw new DocuChefException($"Failed to initialize Excel template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Initialize the template with built-in functions and global variables
    /// </summary>
    private void InitializeTemplate()
    {
        if (_options.RegisterBuiltInFunctions)
        {
            _template.RegisterBuiltIns();
            Logger.Debug("Registered built-in functions for Excel template");
        }

        if (_options.RegisterGlobalVariables)
        {
            RegisterStandardGlobalVariables();
            Logger.Debug("Registered global variables for Excel template");
        }
    }

    /// <summary>
    /// Adds a variable to the template
    /// </summary>
    public override void AddVariable(string name, object value)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        try
        {
            Logger.Debug($"Adding variable '{name}' to Excel template");
            _template.AddVariable(name, value);
            Variables[name] = value;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to add variable '{name}'", ex);
            throw new DocuChefException($"Failed to add variable '{name}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Registers a custom function for cell processing
    /// </summary>
    public void RegisterFunction(string name, Action<IXLCell, object, string[]> function)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (function == null)
            throw new ArgumentNullException(nameof(function));

        try
        {
            // Convert to XLFunctionHandler - which is what XLCustomTemplate.RegisterFunction expects
            XLFunctionHandler handler = (cell, value, parameters) => function(cell, value, parameters);
            _template.RegisterFunction(name, handler);
            Logger.Debug($"Registered function '{name}' for Excel template");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to register function '{name}'", ex);
            throw new DocuChefException($"Failed to register function '{name}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Generates the document from the template
    /// </summary>
    public ExcelDocument Generate()
    {
        ThrowIfDisposed();

        try
        {
            Logger.Debug("Generating Excel document from template");
            _template.Generate();
            _isGenerated = true;

            // Get the workbook from the template
            var workbook = _template.Workbook;
            if (workbook == null)
            {
                throw new DocuChefException("Failed to retrieve workbook from template after generation.");
            }

            Logger.Info("Excel document generated successfully");
            return new ExcelDocument(workbook);
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to generate Excel document", ex);
            throw new DocuChefException($"Failed to generate Excel document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        if (IsDisposed) return;

        if (disposing)
        {
            _template?.Dispose();
            Logger.Debug("Excel template disposed");
        }

        base.Dispose(disposing);
    }
}