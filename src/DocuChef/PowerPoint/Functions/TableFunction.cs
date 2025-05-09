namespace DocuChef.PowerPoint.Functions;

/// <summary>
/// Table-related functions for PowerPoint processing according to PPT syntax guidelines
/// </summary>
internal static class TableFunction
{
    /// <summary>
    /// Creates a PowerPoint function for table handling
    /// </summary>
    public static PowerPointFunction Create()
    {
        return new PowerPointFunction
        {
            Name = "Table",
            Description = "Inserts or updates a table in a PowerPoint shape according to ppt.Table syntax",
            Handler = ProcessTableFunction
        };
    }

    /// <summary>
    /// Process table function: ppt.Table("dataSource", headers: true, startRow: 1, endRow: 10, style: "Medium")
    /// </summary>
    private static object ProcessTableFunction(PowerPointContext context, object value, string[] parameters)
    {
        // Expected parameters:
        // 0: Data source property name
        if (parameters == null || parameters.Length == 0)
        {
            Logger.Warning("Table function called without required data source parameter");
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
        bool headers = true;
        int startRow = 0;
        int endRow = -1; // -1 means all rows
        string style = "Medium"; // Default table style

        // Parse named parameters according to PPT syntax (headers: true, startRow: 1, endRow: 10)
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
                    case "headers":
                        if (bool.TryParse(paramValue, out bool h))
                            headers = h;
                        break;
                    case "startrow":
                        if (int.TryParse(paramValue, out int sr))
                            startRow = sr;
                        break;
                    case "endrow":
                        if (int.TryParse(paramValue, out int er))
                            endRow = er;
                        break;
                    case "style":
                        style = paramValue;
                        break;
                }
            }
        }

        Logger.Debug($"Processing table with parameters - headers:{headers}, startRow:{startRow}, endRow:{endRow}, style:{style}");

        // Begin table processing
        try
        {
            // Get the shape and slide
            var shape = context.Shape;
            var slide = context.Slide;

            // This is a simplified implementation - for a complete implementation, 
            // we would need to convert the data object to a table data structure,
            // and then create or update the table using OpenXML

            // For now, return a placeholder message
            return $"[Table: {dataSource}, Headers: {headers}, StartRow: {startRow}, EndRow: {endRow}, Style: {style}]";

            // Complete implementation would include:
            // 1. Convert data object to table data
            // 2. Create or update table in shape
            // 3. Apply table style
            // 4. Return empty string on success
        }
        catch (Exception ex)
        {
            Logger.Error("Error processing table", ex);
            return $"[Error processing table: {ex.Message}]";
        }
    }
}