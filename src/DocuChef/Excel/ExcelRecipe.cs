using ClosedXML.Excel;
using ClosedXML.Report;
namespace DocuChef.Excel;

public class ExcelRecipe : RecipeBase<IXLWorkbook>
{
    private XLTemplate _template;

    /// <summary>
    /// Creates a new Excel recipe from the specified template
    /// </summary>
    public ExcelRecipe(string templatePath, RecipeOptions options)
        : base(templatePath, options)
    {
        try
        {
            _template = new XLTemplate(templatePath);
            LoggingHelper.LogInformation($"Excel template loaded: {templatePath}");
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error initializing Excel recipe", ex);
            throw;
        }
    }

    /// <summary>
    /// Process the template with the provided data
    /// </summary>
    protected override async Task ProcessTemplateAsync()
    {
        try
        {
            LoggingHelper.LogInformation("Processing Excel template");

            // 데이터를 직접 템플릿에 추가
            // ClosedXML.Report는 이미 dynamic 객체를 잘 처리함
            // 변환 과정을 단순화
            _template.AddVariable(Data);

            // 템플릿 생성
            _template.Generate();

            // 문서 참조 설정
            Document = _template.Workbook;

            // Excel 옵션 적용
            if (Options?.Excel != null)
            {
                ApplyExcelOptions(Document);
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error processing Excel template", ex);
            throw new TemplateException($"Error processing Excel template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Register all named ranges in the template
    /// </summary>
    private void RegisterDefinedNames()
    {
        try
        {
            foreach (var range in _template.Workbook.DefinedNames)
            {
                try
                {
                    // Try to find and use the correct method to register ranges
                    var methodInfo = _template.GetType().GetMethod("RegisterRange",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

                    if (methodInfo != null)
                    {
                        methodInfo.Invoke(_template, new object[] { range.Name });
                        LoggingHelper.LogInformation($"Registered range: {range.Name}");
                        continue;
                    }

                    // Alternative approach if RegisterRange is not available
                    var rangeObj = _template.Workbook.NamedRange(range.Name);
                    if (rangeObj != null)
                    {
                        // Add as a variable instead
                        _template.AddVariable(range.Name, rangeObj);
                        LoggingHelper.LogInformation($"Added range as variable: {range.Name}");
                    }
                }
                catch (Exception ex)
                {
                    LoggingHelper.LogWarning($"Failed to register range {range.Name}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogWarning($"Error registering ranges: {ex.Message}");
        }
    }

    /// <summary>
    /// Convert dictionary data to an anonymous object for ClosedXML.Report
    /// </summary>
    private object ConvertToAnonymousObject(Dictionary<string, object> data)
    {
        // Create a dynamic ExpandoObject for better template binding
        dynamic result = new System.Dynamic.ExpandoObject();
        var resultDict = (IDictionary<string, object>)result;

        // Copy data to ExpandoObject
        foreach (var item in data)
        {
            resultDict[item.Key] = item.Value;
            LoggingHelper.LogInformation($"Added data: {item.Key}");
        }

        return result;
    }

    /// <summary>
    /// Apply Excel-specific options to the workbook
    /// </summary>
    private void ApplyExcelOptions(IXLWorkbook workbook)
    {
        if (workbook == null) return;

        try
        {
            // Apply autofit columns if enabled
            if (Options.Excel.AutoFitColumns)
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Columns().AdjustToContents();
                }
                LoggingHelper.LogInformation("Applied auto-fit columns");
            }

            // Execute plugins if any
            foreach (var plugin in Options.Excel.Plugins)
            {
                plugin.Execute(workbook, Data, Options);
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogWarning($"Failed to apply Excel options: {ex.Message}");
        }
    }

    /// <summary>
    /// Save Excel document to the specified path
    /// </summary>
    protected override async Task SaveDocumentAsync(string outputPath)
    {
        try
        {
            EnsureDirectoryExists(outputPath);

            if (Document == null)
            {
                throw new InvalidOperationException("Document is not initialized.");
            }

            await Task.Run(() => Document.SaveAs(outputPath));
            LoggingHelper.LogInformation($"Document saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error saving Excel document", ex);
            throw new DocuChefException($"Error saving Excel document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Reload document from template
    /// </summary>
    protected override void ReloadDocument()
    {
        try
        {
            _template?.Dispose();
            _template = new XLTemplate(TemplatePath);
            Document = _template.Workbook;
            LoggingHelper.LogInformation($"Excel document reloaded: {TemplatePath}");
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error reloading Excel template", ex);
            throw;
        }
    }

    /// <summary>
    /// Dispose resources
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _template?.Dispose();
        }
        base.Dispose(disposing);
    }
}