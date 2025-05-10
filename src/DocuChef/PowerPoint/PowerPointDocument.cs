using DocumentFormat.OpenXml.Packaging;

namespace DocuChef.PowerPoint;

/// <summary>
/// Represents a generated PowerPoint document
/// </summary>
public class PowerPointDocument : IDisposable
{
    private readonly PresentationDocument _presentationDocument;
    private readonly string _documentPath;
    private bool _isDisposed;
    private bool _isSaved;

    /// <summary>
    /// The underlying OpenXml PresentationDocument instance
    /// </summary>
    public PresentationDocument PresentationDocument => _presentationDocument;

    /// <summary>
    /// The path to the temporary document file
    /// </summary>
    internal string DocumentPath => _documentPath;

    internal PowerPointDocument(PresentationDocument presentationDocument, string documentPath)
    {
        _presentationDocument = presentationDocument ?? throw new ArgumentNullException(nameof(presentationDocument));
        _documentPath = documentPath;
        _isSaved = false;
    }

    /// <summary>
    /// Saves the document to the specified path
    /// </summary>
    public void SaveAs(string filePath)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));

        try
        {
            // Save all document parts
            SaveAllDocumentParts();

            // Dispose the presentation document
            _presentationDocument.Dispose();
            _isSaved = true;
            Logger.Debug($"Presentation document disposed properly");

            // Ensure target directory exists
            filePath.EnsureDirectoryExists();
            Logger.Debug($"Ensured directory exists for: {filePath}");

            // Copy to final location
            File.Copy(_documentPath, filePath, true);
            Logger.Debug($"File copied from {_documentPath} to {filePath}");

            // Verify the output file
            if (!File.Exists(filePath))
            {
                Logger.Error($"Output file not found after copy: {filePath}");
                throw new DocuChefException("Failed to create output file.");
            }

            var fileInfo = new FileInfo(filePath);
            if (fileInfo.Length == 0)
            {
                Logger.Error($"Output file is empty: {filePath}");
                throw new DocuChefException("Output file is empty.");
            }

            Logger.Info($"PowerPoint document saved successfully to {filePath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to save PowerPoint document to {filePath}", ex);
            throw new DocuChefException($"Failed to save PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Saves the document to a stream
    /// </summary>
    public void SaveAs(Stream stream)
    {
        ThrowIfDisposed();

        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        try
        {
            // Save all document parts
            SaveAllDocumentParts();

            // Dispose the presentation document
            _presentationDocument.Dispose();
            _isSaved = true;
            Logger.Debug($"Presentation document disposed properly");

            // Copy to stream
            using var fileStream = new FileStream(_documentPath, FileMode.Open, FileAccess.Read);
            fileStream.CopyTo(stream);
            Logger.Debug($"File copied from {_documentPath} to stream.");

            Logger.Info("PowerPoint document saved successfully to stream");
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to save PowerPoint document to stream", ex);
            throw new DocuChefException($"Failed to save PowerPoint document to stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Saves all document parts to ensure proper document structure
    /// </summary>
    private void SaveAllDocumentParts()
    {
        if (_presentationDocument?.PresentationPart == null)
            return;

        try
        {
            // Save the main presentation part
            _presentationDocument.PresentationPart.Presentation.Save();
            Logger.Debug("Presentation part saved");

            // Save all slide parts
            foreach (var slidePart in _presentationDocument.PresentationPart.SlideParts)
            {
                if (slidePart?.Slide != null)
                {
                    slidePart.Slide.Save();
                }
            }
            Logger.Debug("All slide parts saved");

            // Save all slide master parts
            foreach (var masterPart in _presentationDocument.PresentationPart.SlideMasterParts)
            {
                if (masterPart?.SlideMaster != null)
                {
                    masterPart.SlideMaster.Save();
                }
            }
            Logger.Debug("All slide master parts saved");

            // Save theme part if exists
            if (_presentationDocument.PresentationPart.ThemePart?.Theme != null)
            {
                _presentationDocument.PresentationPart.ThemePart.Theme.Save();
            }

            // Save view properties if exists
            if (_presentationDocument.PresentationPart.ViewPropertiesPart?.ViewProperties != null)
            {
                _presentationDocument.PresentationPart.ViewPropertiesPart.ViewProperties.Save();
            }

            // Save presentation properties if exists
            if (_presentationDocument.PresentationPart.PresentationPropertiesPart?.PresentationProperties != null)
            {
                _presentationDocument.PresentationPart.PresentationPropertiesPart.PresentationProperties.Save();
            }

            // Save the document package
            _presentationDocument.Save();
            Logger.Debug("Document package saved");
        }
        catch (Exception ex)
        {
            Logger.Error("Error saving document parts", ex);
            throw;
        }
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (_isDisposed) return;

        if (disposing)
        {
            try
            {
                if (!_isSaved && _presentationDocument != null)
                {
                    _presentationDocument.Dispose();
                    Logger.Debug("Presentation document disposed during disposal");
                }

                if (!string.IsNullOrEmpty(_documentPath) && File.Exists(_documentPath))
                {
                    File.Delete(_documentPath);
                    Logger.Debug($"Temporary file deleted: {_documentPath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error disposing PowerPoint document resources", ex);
            }
        }

        _isDisposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(PowerPointDocument));
    }
}