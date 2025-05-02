namespace DocuChef;

/// <summary>
/// Base interface for all document recipes, providing common methods for data binding and document generation
/// </summary>
public interface IRecipe : IDisposable, IAsyncDisposable
{
    /// <summary>
    /// Adds data to the template for binding
    /// </summary>
    IRecipe AddData(object data);

    /// <summary>
    /// Generates and saves the document to the specified path
    /// </summary>
    Task SaveAsync(string outputPath);

    /// <summary>
    /// Resets the template data for reuse
    /// </summary>
    void Reset();
}

/// <summary>
/// Generic version of IRecipe allowing access to the underlying document
/// </summary>
public interface IRecipe<TDocument> : IRecipe
{
    /// <summary>
    /// Returns the document object for direct manipulation
    /// </summary>
    TDocument GetDocument();
}

public static class RecipeExtensions
{
    /// <summary>
    /// Seasons your document template with fresh data ingredients.
    /// This culinary-inspired method wraps the standard AddData operation
    /// with more flavorful terminology that matches DocuChef's theming.
    /// </summary>
    /// <param name="recipe">The document recipe to prepare</param>
    /// <param name="data">The data ingredients to mix into your document</param>
    /// <returns>The recipe instance for method chaining</returns>
    public static IRecipe AddIngredients(this IRecipe recipe, object data)
    {
        return recipe.AddData(data);
    }

    /// <summary>
    /// Plates and presents your completed document with professional flair.
    /// This chef-inspired alternative to the standard SaveAsync method
    /// delivers your finished creation to its destination path.
    /// </summary>
    /// <param name="recipe">The prepared document recipe</param>
    /// <param name="outputPath">The serving location for your finished document</param>
    /// <returns>A task representing the asynchronous save operation</returns>
    public static Task ServeAsync(this IRecipe recipe, string outputPath)
    {
        return recipe.SaveAsync(outputPath);
    }
}