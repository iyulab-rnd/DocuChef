namespace DocuChef.PowerPoint.Functions;

/// <summary>
/// Chart-related functions for PowerPoint processing according to PPT syntax guidelines
/// </summary>
internal static class ChartFunction
{
    /// <summary>
    /// Creates a PowerPoint function for chart handling
    /// </summary>
    public static PowerPointFunction Create()
    {
        return new PowerPointFunction
        {
            Name = "Chart",
            Description = "Inserts or updates a chart in a PowerPoint shape according to ppt.Chart syntax",
            Handler = ProcessChartFunction
        };
    }

    /// <summary>
    /// Process chart function: ppt.Chart("dataSource", series: "series", categories: "categories", title: "title")
    /// </summary>
    private static object ProcessChartFunction(PowerPointContext context, object value, string[] parameters)
    {
        // Expected parameters:
        // 0: Data source property name
        if (parameters == null || parameters.Length == 0)
        {
            Logger.Warning("Chart function called without required data source parameter");
            return "[Error: Data source required]";
        }

        string dataSource = parameters[0];
        object dataObject = null;

        // Resolve data source from variables
        dataObject = context.ResolveVariable(dataSource);

        // If data object is null, return error
        if (dataObject == null)
        {
            Logger.Warning($"Data source '{dataSource}' not found");
            return $"[Error: Data source '{dataSource}' not found]";
        }

        // Default parameter values
        string series = null;
        string categories = null;
        string title = null;
        string chartType = "Column"; // Default chart type

        // Parse named parameters according to PPT syntax (series: "series", categories: "categories")
        for (int i = 1; i < parameters.Length; i++)
        {
            string param = parameters[i];

            // Split by first colon for named parameters
            var colonIndex = param.IndexOf(':');
            if (colonIndex > 0)
            {
                string paramName = param.Substring(0, colonIndex).Trim();
                string paramValue = param.Substring(colonIndex + 1).Trim();

                switch (paramName.ToLowerInvariant())
                {
                    case "series":
                        series = paramValue;
                        break;
                    case "categories":
                        categories = paramValue;
                        break;
                    case "title":
                        title = paramValue;
                        break;
                    case "type":
                        chartType = paramValue;
                        break;
                }
            }
        }

        Logger.Debug($"Processing chart with parameters - type:{chartType}, series:{series}, categories:{categories}, title:{title}");

        // Begin chart processing
        try
        {
            // Get the shape and slide
            var shape = context.Shape;
            var slide = context.Slide;

            // This is a simplified implementation - for a complete implementation, 
            // we would need to convert the data object to a chart data structure,
            // and then create or update the chart using OpenXML

            // For now, return a placeholder message
            return $"[Chart: {dataSource}, Type: {chartType}, Series: {series}, Categories: {categories}, Title: {title}]";

            // Complete implementation would include:
            // 1. Convert data object to chart data
            // 2. Create or update chart in shape
            // 3. Update chart type, series, categories, title
            // 4. Return empty string on success
        }
        catch (Exception ex)
        {
            Logger.Error("Error processing chart", ex);
            return $"[Error processing chart: {ex.Message}]";
        }
    }
}