namespace DocuChef.Exceptions;

/// <summary>
/// Exception thrown when there is an error processing a template
/// </summary>
public class TemplateException : DocuChefException
{
    /// <summary>
    /// Creates a new template exception with the specified message
    /// </summary>
    public TemplateException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new template exception with the specified message and inner exception
    /// </summary>
    public TemplateException(string message, Exception innerException) : base(message, innerException)
    {
    }
}