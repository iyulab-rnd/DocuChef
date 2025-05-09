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
            // Make sure the document is saved
            _presentationDocument.Save();
            Logger.Debug($"Presentation document saved to temporary file: {_documentPath}");

            // Dispose to release the file handle
            _presentationDocument.Dispose();
            Logger.Debug("Presentation document disposed.");

            // Create necessary directories and copy the file
            filePath.EnsureDirectoryExists();
            Logger.Debug($"Ensured directory exists for: {filePath}");

            File.Copy(_documentPath, filePath, true);
            Logger.Debug($"File copied from {_documentPath} to {filePath}");

            // Verify output file
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
            // Make sure the document is saved
            _presentationDocument.Save();
            Logger.Debug($"Presentation document saved to temporary file: {_documentPath}");

            // Dispose to release the file handle
            _presentationDocument.Dispose();
            Logger.Debug("Presentation document disposed.");

            // Copy the temporary file to the stream
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
                _presentationDocument?.Dispose();
                Logger.Debug("Presentation document disposed");

                // Delete the temporary file
                if (!string.IsNullOrEmpty(_documentPath) && File.Exists(_documentPath))
                {
                    File.Delete(_documentPath);
                    Logger.Debug($"Temporary file deleted: {_documentPath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error disposing PowerPoint document resources", ex);
                // Ignore disposal errors
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