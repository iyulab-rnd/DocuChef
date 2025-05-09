using DocuChef.PowerPoint.Functions;

namespace DocuChef.PowerPoint;

/// <summary>
/// Built-in functions for PowerPoint processing according to PPT syntax guidelines
/// </summary>
internal static class PowerPointFunctions
{
    /// <summary>
    /// Register all built-in PowerPoint functions
    /// </summary>
    public static void RegisterBuiltInFunctions(PowerPointRecipe recipe)
    {
        // Register image function - ppt.Image("imageProperty", width: 300, height: 200)
        recipe.RegisterFunction(ImageFunction.Create());

        // Register chart function - ppt.Chart("dataSource", series: "series", categories: "categories")
        recipe.RegisterFunction(ChartFunction.Create());

        // Register table function - ppt.Table("dataSource", headers: true, startRow: 1, endRow: 10)
        recipe.RegisterFunction(TableFunction.Create());
    }
}