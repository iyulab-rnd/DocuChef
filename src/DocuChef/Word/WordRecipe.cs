using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace DocuChef.Word;

/// <summary>
/// Word template engine implementation using DollarSignEngine
/// </summary>
public partial class WordRecipe : RecipeBase<Document>
{
    private WordprocessingDocument? _wordDoc;
    private static readonly Regex SectionRegex = new(@"<!--#begin:(\w+):(\w+)-->(.*?)<!--#end:\1:\2-->", RegexOptions.Compiled | RegexOptions.Singleline);
    private static readonly Regex IfRegex = new(@"<!--#if:(\w+)-->(.*?)(?:<!--#else:\1-->(.*?))?<!--#endif:\1-->", RegexOptions.Compiled | RegexOptions.Singleline);

    /// <summary>
    /// Callback executed when the document is created
    /// </summary>
    public Action<Document>? OnDocumentCreated { get; set; }

    /// <summary>
    /// Creates a new Word recipe from the specified template
    /// </summary>
    public WordRecipe(string templatePath, RecipeOptions options)
        : base(templatePath, options)
    {
        ReloadDocument();
    }

    /// <summary>
    /// Processes the Word template with the provided data
    /// </summary>
    protected override async Task ProcessTemplateAsync()
    {
        try
        {
            if (Document == null || _wordDoc == null)
            {
                throw new InvalidOperationException("Document is not initialized.");
            }

            LoggingHelper.LogInformation("Processing main document content");
            await ProcessMainDocumentAsync();

            if (Options.Word.ProcessHeadersAndFooters)
            {
                LoggingHelper.LogInformation("Processing headers and footers");
                await ProcessHeadersAndFootersAsync();
            }

            LoggingHelper.LogInformation("Processing sections (collection data)");
            await ProcessSectionsAsync();

            LoggingHelper.LogInformation("Executing plugins");
            foreach (var plugin in Options.Word.Plugins)
            {
                plugin.Execute(Document, Data, Options);
            }

            if (Options.Word.UpdateFieldsAfterBinding)
            {
                LoggingHelper.LogInformation("Updating fields");
                await UpdateFieldsAsync();
            }

            if (Options.Word.UpdateTableOfContents)
            {
                LoggingHelper.LogInformation("Updating table of contents");
                await UpdateTableOfContentsAsync();
            }

            if (OnDocumentCreated != null && Document != null)
            {
                LoggingHelper.LogInformation("Executing document callback");
                OnDocumentCreated(Document);
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error processing Word template", ex);
            throw new DocuChefException($"Error processing Word template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Saves the Word document to the specified path
    /// </summary>
    protected override async Task SaveDocumentAsync(string outputPath)
    {
        try
        {
            EnsureDirectoryExists(outputPath);

            if (_wordDoc == null)
            {
                throw new InvalidOperationException("Document is not initialized.");
            }

            // Save as a copy following the same pattern as in PowerPoint
            await Task.Run(() => {
                // We need to use the OpenXML Save method to write to the output file
                // Create a temporary file first
                var tempPath = Path.GetTempFileName() + ".docx";

                // Save current document in memory
                _wordDoc.Save(); // Save any changes in memory

                // Create a new document at the target location
                using (var sourceDoc = _wordDoc)
                using (var destDoc = WordprocessingDocument.Create(tempPath, WordprocessingDocumentType.Document))
                {
                    // Copy all parts from source to destination
                    foreach (var part in sourceDoc.GetAllParts())
                    {
                        destDoc.AddPart(part);
                    }

                    // Save the new document
                    destDoc.Save();
                }

                // Move the file to the target location
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
                File.Move(tempPath, outputPath);

                // Reopen the original document
                _wordDoc = WordprocessingDocument.Open(TemplatePath, false);
                Document = _wordDoc.MainDocumentPart?.Document;

                if (Document == null)
                {
                    throw new InvalidOperationException("Failed to reload document content.");
                }

                LoggingHelper.LogInformation($"Document saved to: {outputPath}");
            });
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error saving Word document", ex);
            throw new DocuChefException($"Error saving Word document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Reloads the Word document from the template
    /// </summary>
    protected override void ReloadDocument()
    {
        try
        {
            if (_wordDoc != null)
            {
                _wordDoc.Dispose();
            }

            _wordDoc = WordprocessingDocument.Open(TemplatePath, false);
            Document = _wordDoc.MainDocumentPart?.Document;

            if (Document == null)
            {
                throw new InvalidOperationException("Failed to load document content.");
            }

            LoggingHelper.LogInformation($"Word document loaded: {TemplatePath}");
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error loading Word template", ex);
            throw new DocuChefException($"Error loading Word template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Dispose pattern implementation
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_wordDoc != null)
            {
                _wordDoc.Dispose();
                _wordDoc = null;
            }
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// Async dispose implementation
    /// </summary>
    protected override async ValueTask DisposeAsyncCore()
    {
        if (_wordDoc != null)
        {
            await Task.Run(() => _wordDoc.Dispose());
            _wordDoc = null;
        }

        await base.DisposeAsyncCore();
    }
}