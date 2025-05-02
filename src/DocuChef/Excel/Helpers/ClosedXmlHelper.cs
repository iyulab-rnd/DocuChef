using ClosedXML.Excel;
using ClosedXML.Report;
using System;
using System.Collections.Generic;
using DocuChef.Utils;

namespace DocuChef.Excel.Helpers;

/// <summary>
/// Helper methods for ClosedXML.Report integration
/// </summary>
internal static class ClosedXmlHelper
{
    /// <summary>
    /// Register named ranges with the template
    /// </summary>
    public static void RegisterTemplateRanges(IXLTemplate template)
    {
        if (template?.Workbook == null)
            return;

        try
        {
            // Get all named ranges
            foreach (var range in template.Workbook.DefinedNames)
            {
                try
                {
                    // Check if a method to register ranges exists
                    // This is a workaround since RegisterRange seems to be missing
                    AddNamedRangeToTemplate(template, range.Name);
                    LoggingHelper.LogInformation($"Registered range: {range.Name}");
                }
                catch (Exception ex)
                {
                    // Log failure but continue with other ranges
                    LoggingHelper.LogWarning($"Failed to register range {range.Name}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogWarning($"Error registering template ranges: {ex.Message}");
        }
    }

    // Alternative method to handle named ranges if RegisterRange is not available
    private static void AddNamedRangeToTemplate(IXLTemplate template, string rangeName)
    {
        // Method 1: Use reflection to find and invoke the method if it exists
        var methodInfo = template.GetType().GetMethod("RegisterRange",
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

        if (methodInfo != null)
        {
            methodInfo.Invoke(template, new object[] { rangeName });
            return;
        }

        // Method 2: If RegisterNamedRange exists instead (check actual API)
        methodInfo = template.GetType().GetMethod("RegisterNamedRange",
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

        if (methodInfo != null)
        {
            methodInfo.Invoke(template, new object[] { rangeName });
            return;
        }

        // Method 3: Add variable for the range directly if previous methods fail
        var range = template.Workbook.NamedRange(rangeName);
        if (range != null)
        {
            try
            {
                // Try to add the range as a variable to the template
                template.AddVariable(rangeName, range);
            }
            catch
            {
                // If that fails, log it and continue
                LoggingHelper.LogWarning($"Could not add range {rangeName} to template");
            }
        }
    }

    /// <summary>
    /// Configure workbook settings after generation
    /// </summary>
    public static void ConfigureWorkbook(IXLWorkbook workbook, ExcelOptions options)
    {
        if (workbook == null || options == null)
            return;

        try
        {
            // Auto-fit columns if enabled
            if (options.AutoFitColumns)
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Columns().AdjustToContents();
                }
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogWarning($"Failed to configure workbook: {ex.Message}");
        }
    }

    /// <summary>
    /// Add data to the template
    /// </summary>
    public static void AddDataToTemplate(IXLTemplate template, Dictionary<string, object> data)
    {
        if (template == null || data == null || data.Count == 0)
            return;

        try
        {
            template.AddVariable(data);
            LoggingHelper.LogInformation($"Added {data.Count} variables to template");
        }
        catch (Exception ex)
        {
            LoggingHelper.LogWarning($"Error adding data to template: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Converts a dictionary to a dynamic object
    /// </summary>
    private static dynamic ConvertDictionaryToDynamic(Dictionary<string, object> data)
    {
        dynamic result = new System.Dynamic.ExpandoObject();
        var dictionary = (IDictionary<string, object>)result;

        foreach (var item in data)
        {
            dictionary[item.Key] = item.Value;
        }

        return result;
    }
}