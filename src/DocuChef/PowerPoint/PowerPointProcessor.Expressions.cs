namespace DocuChef.PowerPoint;

/// <summary>
/// Partial class for PowerPointProcessor - Expression evaluation methods
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Evaluate expression with provided variables
    /// </summary>
    public object EvaluateCompleteExpression(string expression, Dictionary<string, object> variables)
    {
        // Use DollarSignEngine adapter to evaluate
        try
        {
            return _expressionEvaluator.Evaluate(expression, variables);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// Evaluate expression using DollarSignEngine
    /// </summary>
    internal object EvaluateCompleteExpression(string expression)
    {
        // Prepare variables dictionary
        var variables = PrepareVariables();
        return EvaluateCompleteExpression(expression, variables);
    }

    /// <summary>
    /// Resolve variable value from context
    /// </summary>
    private object ResolveVariableValue(string name)
    {
        // Check direct variable
        if (_context.Variables.TryGetValue(name, out var value))
            return value;

        // Check global variable
        if (_context.GlobalVariables.TryGetValue(name, out var factory))
            return factory();

        // Check property path
        if (name.Contains('.'))
        {
            var parts = name.Split('.');
            if (_context.Variables.TryGetValue(parts[0], out var obj))
            {
                for (int i = 1; i < parts.Length && obj != null; i++)
                {
                    var property = obj.GetType().GetProperty(parts[i]);
                    obj = property?.GetValue(obj);
                }
                return obj;
            }
        }

        return null;
    }

    /// <summary>
    /// Convert an object to a list of objects for array processing
    /// </summary>
    private List<object> ConvertToList(object obj)
    {
        if (obj == null)
            return null;

        if (obj is IList list)
        {
            return list.Cast<object>().ToList();
        }
        else if (obj is IEnumerable enumerable)
        {
            return enumerable.Cast<object>().ToList();
        }

        // Not a collection, return a single-item list
        return new List<object> { obj };
    }
}