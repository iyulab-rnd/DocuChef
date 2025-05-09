using System.Text.RegularExpressions;
using DollarSignEngine;

namespace DocuChef.PowerPoint.DollarSignEngine;

/// <summary>
/// Adapter for DollarSignEngine to handle expressions in PowerPoint templates
/// </summary>
internal class ExpressionEvaluator
{
    private readonly DollarSignOption _options;

    /// <summary>
    /// Initializes a new instance of the ExpressionEvaluator
    /// </summary>
    public ExpressionEvaluator()
    {
        _options = new DollarSignOption
        {
            SupportDollarSignSyntax = true,  // Enable ${variable} syntax as per PPT syntax guidelines
            ThrowOnMissingParameter = false, // Don't throw on missing parameters, show placeholder instead
            EnableDebugLogging = false,      // Disable debug logging by default
            PreferCallbackResolution = true, // Prefer callback resolution for special functions
            VariableResolver = HandleSpecialFunctions // Custom resolver for PowerPoint functions
        };
    }

    /// <summary>
    /// Evaluates an expression synchronously
    /// </summary>
    public object Evaluate(string expression, Dictionary<string, object> variables)
    {
        return EvaluateAsync(expression, variables).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Evaluates an expression asynchronously
    /// </summary>
    public async Task<object> EvaluateAsync(string expression, Dictionary<string, object> variables)
    {
        try
        {
            Logger.Debug($"Evaluating expression: {expression}");

            // If expression is a PowerPoint special function (ppt.)
            if (expression.StartsWith("ppt."))
            {
                return await HandlePptFunctionAsync(expression, variables);
            }

            // If expression already contains ${...}, just evaluate it directly
            if (expression.Contains("${"))
            {
                Logger.Debug($"Evaluating text with embedded variables: {expression}");
                var result = await DollarSign.EvalAsync(expression, variables, _options);
                Logger.Debug($"Expression result: {result}");
                return result;
            }

            // If expression doesn't start with ${...}, wrap it for evaluation
            if (!expression.StartsWith("${"))
            {
                expression = "${" + expression + "}";
            }

            // Evaluate using DollarSignEngine
            var evalResult = await DollarSign.EvalAsync(expression, variables, _options);
            Logger.Debug($"Expression result: {evalResult}");
            return evalResult;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{expression}'", ex);
            throw new DocuChefException($"Error evaluating expression '{expression}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Custom resolver for PowerPoint special functions
    /// </summary>
    private object HandleSpecialFunctions(string expression, object parameters)
    {
        // If the expression starts with "ppt.", handle it as a PowerPoint function
        if (expression.StartsWith("ppt."))
        {
            // Convert parameters to dictionary if needed
            Dictionary<string, object> variables;
            if (parameters is Dictionary<string, object> dict)
            {
                variables = dict;
            }
            else
            {
                variables = new Dictionary<string, object>();
                if (parameters != null)
                {
                    var props = parameters.GetType().GetProperties();
                    foreach (var prop in props)
                    {
                        if (prop.CanRead)
                        {
                            try
                            {
                                var value = prop.GetValue(parameters);
                                variables[prop.Name] = value;
                            }
                            catch
                            {
                                // Skip properties that throw exceptions
                            }
                        }
                    }
                }
            }

            // Handle the PowerPoint function
            var task = HandlePptFunctionAsync(expression, variables);
            return task.GetAwaiter().GetResult();
        }

        // Return null to let DollarSignEngine handle standard expressions
        return null;
    }

    /// <summary>
    /// Handle PowerPoint special functions (ppt.Image, ppt.Chart, ppt.Table)
    /// </summary>
    private async Task<object> HandlePptFunctionAsync(string expression, Dictionary<string, object> variables)
    {
        Logger.Debug($"Handling PowerPoint function: {expression}");

        // Parse function expression: ppt.Function("arg", param1: value1, param2: value2)
        var match = Regex.Match(expression, @"ppt\.(\w+)\((.+)\)");
        if (!match.Success)
        {
            Logger.Warning($"Invalid PowerPoint function format: {expression}");
            return $"[Invalid function: {expression}]";
        }

        string functionName = match.Groups[1].Value;
        string argsString = match.Groups[2].Value;

        Logger.Debug($"Function: {functionName}, Args: {argsString}");

        // Parse arguments with proper handling of quoted strings and named parameters
        var args = ParseFunctionArguments(argsString);

        // Resolve arguments using DollarSign engine if they contain expressions
        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            // If the argument contains ${...} or is a named parameter like param:${value}
            if (arg.Contains("${") || (arg.Contains(":") && arg.Split(':', 2)[1].Contains("${")))
            {
                // For named parameters (param:value), only evaluate the value part
                if (arg.Contains(":"))
                {
                    var parts = arg.Split(':', 2);
                    string paramName = parts[0].Trim();
                    string paramValue = parts[1].Trim();

                    // Evaluate the parameter value
                    var resolvedValue = await DollarSign.EvalAsync(paramValue, variables, _options);
                    args[i] = $"{paramName}: {resolvedValue}";
                }
                else
                {
                    // Direct evaluation of the argument
                    var resolvedValue = await DollarSign.EvalAsync(arg, variables, _options);
                    args[i] = resolvedValue?.ToString() ?? string.Empty;
                }
            }
        }

        // Return a placeholder for now - actual implementation will be done by PowerPointFunctions
        return $"[ppt.{functionName} with resolved args: {string.Join(", ", args)}]";
    }

    /// <summary>
    /// Parse function arguments with improved handling of quoted strings and named parameters
    /// </summary>
    private string[] ParseFunctionArguments(string argsString)
    {
        if (string.IsNullOrEmpty(argsString))
            return Array.Empty<string>();

        var args = new List<string>();
        bool inQuotes = false;
        int start = 0;
        int parenthesesDepth = 0;

        for (int i = 0; i < argsString.Length; i++)
        {
            char c = argsString[i];

            if (c == '"')
            {
                // Handle escaped quotes
                if (i > 0 && argsString[i - 1] == '\\')
                {
                    continue;
                }
                inQuotes = !inQuotes;
            }
            else if (c == '(' && !inQuotes)
            {
                parenthesesDepth++;
            }
            else if (c == ')' && !inQuotes)
            {
                parenthesesDepth--;
            }
            else if (c == ',' && !inQuotes && parenthesesDepth == 0)
            {
                // End of argument
                args.Add(argsString.Substring(start, i - start).Trim());
                start = i + 1;
            }
        }

        // Add the last argument
        if (start < argsString.Length)
        {
            args.Add(argsString.Substring(start).Trim());
        }

        // Clean up quotes and handle named parameters
        for (int i = 0; i < args.Count; i++)
        {
            string arg = args[i].Trim();

            // Handle named parameters (param: value)
            if (arg.Contains(":") && !inQuotes)
            {
                var parts = arg.Split(new[] { ':' }, 2);
                string paramName = parts[0].Trim();
                string paramValue = parts[1].Trim();

                // If the parameter value is quoted, remove the quotes
                if (paramValue.StartsWith("\"") && paramValue.EndsWith("\"") && paramValue.Length > 1)
                {
                    paramValue = paramValue.Substring(1, paramValue.Length - 2)
                        .Replace("\\\"", "\"")
                        .Replace("\\\\", "\\")
                        .Replace("\\n", "\n")
                        .Replace("\\r", "\r");
                }

                arg = $"{paramName}: {paramValue}";
            }
            // Handle regular quoted strings
            else if (arg.StartsWith("\"") && arg.EndsWith("\"") && arg.Length > 1)
            {
                arg = arg.Substring(1, arg.Length - 2)
                    .Replace("\\\"", "\"")
                    .Replace("\\\\", "\\")
                    .Replace("\\n", "\n")
                    .Replace("\\r", "\r");
            }

            args[i] = arg;
        }

        return args.ToArray();
    }
}