namespace DocuChef.Exceptions;

/// <summary>
/// Base exception for all DocuChef errors
/// </summary>
public class DocuChefException : Exception
{
    /// <summary>
    /// Creates a new DocuChef exception with the specified message
    /// </summary>
    public DocuChefException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new DocuChef exception with the specified message and inner exception
    /// </summary>
    public DocuChefException(string message, Exception innerException) : base(message, innerException)
    {
    }
}