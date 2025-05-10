using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint;

/// <summary>
/// Text processing methods for PowerPointProcessor
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Process expressions in text
    /// </summary>
    private string ProcessExpressions(string text)
    {
        if (!ExpressionProcessor.ContainsExpressions(text))
            return text;

        var variables = PrepareVariables();
        return ExpressionProcessor.ProcessExpressions(text, this, variables);
    }

    /// <summary>
    /// Parse function parameters
    /// </summary>
    private string[] ParseFunctionParameters(string parametersString)
    {
        return ExpressionProcessor.ParseFunctionParameters(parametersString);
    }
}