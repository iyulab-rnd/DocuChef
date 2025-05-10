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
        return "TBD";
    }
}