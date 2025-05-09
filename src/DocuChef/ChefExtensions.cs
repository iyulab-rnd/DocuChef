using DocuChef.Excel;
using DocuChef.PowerPoint;

namespace DocuChef;

/// <summary>
/// Provides cooking-themed extension methods for DocuChef
/// </summary>
public static class ChefExtensions
{
    /// <summary>
    /// Loads a template as a recipe
    /// </summary>
    public static IRecipe LoadRecipe(this Chef chef, string templatePath)
    {
        return chef.LoadTemplate(templatePath);
    }

    /// <summary>
    /// Loads an Excel template as a recipe
    /// </summary>
    public static ExcelRecipe LoadExcelRecipe(this Chef chef, string templatePath, ExcelOptions options = null)
    {
        return chef.LoadExcelTemplate(templatePath, options);
    }

    /// <summary>
    /// Loads an Excel template from a stream as a recipe
    /// </summary>
    public static ExcelRecipe LoadExcelRecipe(this Chef chef, Stream templateStream, ExcelOptions options = null)
    {
        return chef.LoadExcelTemplate(templateStream, options);
    }

    /// <summary>
    /// Loads a PowerPoint template as a recipe
    /// </summary>
    public static PowerPointRecipe LoadPresentationRecipe(this Chef chef, string templatePath, PowerPointOptions options = null)
    {
        return chef.LoadPowerPointTemplate(templatePath, options);
    }

    /// <summary>
    /// Loads a PowerPoint template from a stream as a recipe
    /// </summary>
    public static PowerPointRecipe LoadPresentationRecipe(this Chef chef, Stream templateStream, PowerPointOptions options = null)
    {
        return chef.LoadPowerPointTemplate(templateStream, options);
    }
}

/// <summary>
/// Provides cooking-themed extension methods for recipes
/// </summary>
public static class RecipeExtensions
{
    /// <summary>
    /// Adds an ingredient (variable) to the recipe
    /// </summary>
    public static T AddIngredient<T>(this T recipe, string name, object value) where T : IRecipe
    {
        recipe.AddVariable(name, value);
        return recipe;
    }

    /// <summary>
    /// Adds ingredients (variables) from an object to the recipe
    /// </summary>
    public static T AddIngredients<T>(this T recipe, object data) where T : IRecipe
    {
        recipe.AddVariable(data);
        return recipe;
    }

    /// <summary>
    /// Clears all ingredients (variables) from the recipe
    /// </summary>
    public static T ClearIngredients<T>(this T recipe) where T : IRecipe
    {
        recipe.ClearVariables();
        return recipe;
    }

    /// <summary>
    /// Registers a cooking technique (function) for Excel recipes
    /// </summary>
    public static ExcelRecipe RegisterTechnique(this ExcelRecipe recipe, string name, Action<ClosedXML.Excel.IXLCell, object, string[]> function)
    {
        recipe.RegisterFunction(name, function);
        return recipe;
    }

    /// <summary>
    /// Registers a cooking technique (function) for PowerPoint recipes
    /// </summary>
    public static PowerPointRecipe RegisterTechnique(this PowerPointRecipe recipe, PowerPointFunction function)
    {
        recipe.RegisterFunction(function);
        return recipe;
    }

    /// <summary>
    /// Cooks (generates) an Excel recipe
    /// </summary>
    public static ExcelDocument Cook(this ExcelRecipe recipe)
    {
        return recipe.Generate();
    }

    /// <summary>
    /// Cooks (generates) a PowerPoint recipe
    /// </summary>
    public static PowerPointDocument Cook(this PowerPointRecipe recipe)
    {
        return recipe.Generate();
    }
}

/// <summary>
/// Provides cooking-themed extension methods for documents
/// </summary>
public static class DishExtensions
{
    /// <summary>
    /// Serves (saves) a document to a file
    /// </summary>
    public static void Serve<T>(this T document, string filePath) where T : class
    {
        if (document is ExcelDocument excelDoc)
            excelDoc.SaveAs(filePath);
        else if (document is PowerPointDocument powerPointDoc)
            powerPointDoc.SaveAs(filePath);
        else
            throw new InvalidOperationException($"Document type {typeof(T).Name} is not supported");
    }

    /// <summary>
    /// Serves (saves) a document to a stream
    /// </summary>
    public static void Serve<T>(this T document, Stream stream) where T : class
    {
        if (document is ExcelDocument excelDoc)
            excelDoc.SaveAs(stream);
        else if (document is PowerPointDocument powerPointDoc)
            powerPointDoc.SaveAs(stream);
        else
            throw new InvalidOperationException($"Document type {typeof(T).Name} is not supported");
    }
}