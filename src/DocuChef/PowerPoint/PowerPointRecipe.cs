using DocuChef.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace DocuChef.PowerPoint;

/// <summary>
/// PowerPoint template engine implementation using DollarSignEngine
/// </summary>
public partial class PowerPointRecipe : RecipeBase<PresentationDocument>
{
    private static readonly Regex SlideNoteRegex = new(@"@(?<directive>\w+):\s*(?<value>[^@\r\n]+)", RegexOptions.Compiled);

    /// <summary>
    /// Callback executed when the presentation is created
    /// </summary>
    public Action<PresentationDocument>? OnPresentationCreated { get; set; }

    /// <summary>
    /// Creates a new PowerPoint recipe from the specified template
    /// </summary>
    public PowerPointRecipe(string templatePath, RecipeOptions options)
        : base(templatePath, options)
    {
        ReloadDocument();
    }

    /// <summary>
    /// Processes the PowerPoint template with the provided data
    /// </summary>
    protected override async Task ProcessTemplateAsync()
    {
        try
        {
            if (Document == null)
            {
                throw new InvalidOperationException("Document is not initialized.");
            }

            // Process all slides
            LoggingHelper.LogInformation("Processing slides");
            await ProcessSlidesAsync();

            // Execute plugins
            LoggingHelper.LogInformation("Executing plugins");
            foreach (var plugin in Options.PowerPoint.Plugins)
            {
                plugin.Execute(Document, Data, Options);
            }

            // Update slide numbers if required
            if (Options.PowerPoint.UpdateSlideNumbers)
            {
                LoggingHelper.LogInformation("Updating slide numbers");
                await UpdateSlideNumbersAsync();
            }

            // Custom modifications via callback
            if (OnPresentationCreated != null)
            {
                LoggingHelper.LogInformation("Executing presentation callback");
                OnPresentationCreated(Document);
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error processing PowerPoint template", ex);
            throw new DocuChefException($"Error processing PowerPoint template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Saves the PowerPoint document to the specified path
    /// </summary>
    /// <summary>
    /// Saves the PowerPoint document to the specified path
    /// </summary>
    protected override async Task SaveDocumentAsync(string outputPath)
    {
        try
        {
            EnsureDirectoryExists(outputPath);

            if (Document == null)
            {
                throw new InvalidOperationException("Document is not initialized.");
            }

            // Save as a copy following the same pattern as in WordRecipe class
            await Task.Run(() => {
                // We need to use the OpenXML Save method to write to the output file
                // Create a temporary file first
                var tempPath = Path.GetTempFileName() + ".pptx";

                // Save current document to temp file
                Document.Save(); // Save any changes in memory

                // Create a new document at the target location
                using (var sourceDoc = Document)
                using (var destDoc = PresentationDocument.Create(tempPath, PresentationDocumentType.Presentation))
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
                Document = PresentationDocument.Open(TemplatePath, false);

                LoggingHelper.LogInformation($"Document saved to: {outputPath}");
            });
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error saving PowerPoint document", ex);
            throw new DocuChefException($"Error saving PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Reloads the PowerPoint document from the template
    /// </summary>
    protected override void ReloadDocument()
    {
        try
        {
            if (Document != null)
            {
                Document.Dispose();
                Document = null;
            }

            // Use OpenSettings to handle package constraints
            var openSettings = new OpenSettings
            {
                AutoSave = false
            };

            Document = PresentationDocument.Open(TemplatePath, false, openSettings);
            LoggingHelper.LogInformation($"PowerPoint document loaded: {TemplatePath}");

            // Verify that document has required parts
            if (Document.PresentationPart == null)
            {
                throw new InvalidOperationException("Presentation part is missing from the document.");
            }

            if (Document.PresentationPart.Presentation == null)
            {
                throw new InvalidOperationException("Presentation is missing from the presentation part.");
            }
        }
        catch (Exception ex)
        {
            LoggingHelper.LogError("Error loading PowerPoint template", ex);
            throw new DocuChefException($"Error loading PowerPoint template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Async dispose implementation
    /// </summary>
    protected override async ValueTask DisposeAsyncCore()
    {
        if (Document != null)
        {
            await Task.Run(() => {
                try
                {
                    Document.Dispose();
                }
                catch (Exception ex)
                {
                    LoggingHelper.LogWarning($"Error during document disposal: {ex.Message}");
                }
                Document = null;
            });
        }

        await base.DisposeAsyncCore();
    }

    /// <summary>
    /// Dispose pattern implementation
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        if (disposing && Document != null)
        {
            try
            {
                Document.Dispose();
            }
            catch (Exception ex)
            {
                LoggingHelper.LogWarning($"Error during document disposal: {ex.Message}");
            }
            Document = null;
        }

        base.Dispose(disposing);
    }
}