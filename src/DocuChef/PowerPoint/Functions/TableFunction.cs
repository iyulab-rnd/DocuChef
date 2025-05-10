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
        return "TBD";
    }
}