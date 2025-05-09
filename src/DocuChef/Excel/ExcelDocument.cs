using ClosedXML.Excel;

namespace DocuChef.Excel;

/// <summary>
/// Represents a generated Excel document
/// </summary>
public class ExcelDocument : IDisposable
{
    private readonly IXLWorkbook _workbook;
    private bool _isDisposed;

    /// <summary>
    /// The underlying XLWorkbook instance
    /// </summary>
    public IXLWorkbook Workbook => _workbook;

    internal ExcelDocument(IXLWorkbook workbook)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
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
            // Ensure directory exists
            filePath.EnsureDirectoryExists();

            _workbook.SaveAs(filePath);
            Logger.Info($"Excel document saved to {filePath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to save Excel document to {filePath}", ex);
            throw new DocuChefException($"Failed to save Excel document: {ex.Message}", ex);
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
            _workbook.SaveAs(stream);
            Logger.Info("Excel document saved to stream");
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to save Excel document to stream", ex);
            throw new DocuChefException($"Failed to save Excel document to stream: {ex.Message}", ex);
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
            _workbook?.Dispose();
            Logger.Debug("Excel document disposed");
        }

        _isDisposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(ExcelDocument));
    }
}