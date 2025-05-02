namespace DocuChef;

/// <summary>
/// Base implementation for all document recipes
/// </summary>
public abstract class RecipeBase<TDocument> : IRecipe<TDocument>, IAsyncDisposable
    where TDocument : class
{
    protected readonly string TemplatePath;
    protected readonly RecipeOptions Options;
    protected Dictionary<string, object> Data;
    protected TDocument? Document;
    private bool _disposed;

    /// <summary>
    /// Creates a new recipe from the specified template
    /// </summary>
    protected RecipeBase(string templatePath, RecipeOptions options)
    {
        ArgumentNullException.ThrowIfNull(templatePath);

        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"Template file not found: {templatePath}");

        TemplatePath = templatePath;
        Options = options ?? new RecipeOptions();
        Data = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // Configure logging
        if (Options.LogCallback != null)
        {
            LoggingHelper.SetLogCallback(Options.LogCallback);
        }
    }

    /// <summary>
    /// Adds data to the template
    /// </summary>
    public virtual IRecipe AddData(object data)
    {
        ArgumentNullException.ThrowIfNull(data);

        // Merge data into the dictionary
        var newData = DataConverter.ObjectToDictionary(data);
        foreach (var item in newData)
        {
            Data[item.Key] = item.Value;
        }

        return this;
    }

    /// <summary>
    /// Generates and saves the document
    /// </summary>
    public async Task SaveAsync(string outputPath)
    {
        ArgumentNullException.ThrowIfNull(outputPath);

        try
        {
            LoggingHelper.LogInformation($"Processing template: {TemplatePath}");
            await ProcessTemplateAsync();

            LoggingHelper.LogInformation($"Saving document to: {outputPath}");
            await SaveDocumentAsync(outputPath);

            LoggingHelper.LogInformation("Document generation completed successfully");
        }
        catch (Exception ex) when (ex is not DocuChefException)
        {
            LoggingHelper.LogError("Error processing template", ex);
            throw new DocuChefException($"Error processing template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Returns the document for direct manipulation
    /// </summary>
    public TDocument GetDocument()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (Document == null)
            throw new InvalidOperationException("Document has not been initialized.");

        return Document;
    }

    /// <summary>
    /// Resets the template data for reuse
    /// </summary>
    public void Reset()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        LoggingHelper.LogInformation("Resetting template data");
        Data.Clear();
        ReloadDocument();
    }

    /// <summary>
    /// Template method for document processing
    /// </summary>
    protected abstract Task ProcessTemplateAsync();

    /// <summary>
    /// Template method for saving the document
    /// </summary>
    protected abstract Task SaveDocumentAsync(string outputPath);

    /// <summary>
    /// Template method for reloading the document
    /// </summary>
    protected abstract void ReloadDocument();

    /// <summary>
    /// Ensures the output directory exists
    /// </summary>
    protected void EnsureDirectoryExists(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    /// <summary>
    /// Releases all resources
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        await DisposeAsyncCore().ConfigureAwait(false);

        Dispose(disposing: false);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Releases all managed resources
    /// </summary>
    protected virtual async ValueTask DisposeAsyncCore()
    {
        if (Document is IAsyncDisposable asyncDisposable)
        {
            await asyncDisposable.DisposeAsync().ConfigureAwait(false);
        }
        else if (Document is IDisposable disposable)
        {
            disposable.Dispose();
        }

        Document = null;
    }

    /// <summary>
    /// Releases all resources
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Dispose pattern implementation
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // Release managed resources
                if (Document is IDisposable disposable)
                {
                    disposable.Dispose();
                }

                Document = null;
            }

            _disposed = true;
        }
    }

    /// <summary>
    /// Finalizer
    /// </summary>
    ~RecipeBase()
    {
        Dispose(disposing: false);
    }
}