namespace DocuChef.PowerPoint;

/// <summary>
/// Represents a custom function for PowerPoint processing
/// </summary>
public class PowerPointFunction
{
    /// <summary>
    /// Function name
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Function handler
    /// </summary>
    public Func<PowerPointContext, object, string[], object> Handler { get; set; }

    /// <summary>
    /// Function description
    /// </summary>
    public string Description { get; set; }

    /// <summary>
    /// Creates a new PowerPoint function
    /// </summary>
    public PowerPointFunction() { }

    /// <summary>
    /// Creates a new PowerPoint function with the specified properties
    /// </summary>
    public PowerPointFunction(string name, string description, Func<PowerPointContext, object, string[], object> handler)
    {
        Name = name;
        Description = description;
        Handler = handler;
    }

    /// <summary>
    /// Execute the function
    /// </summary>
    public object Execute(PowerPointContext context, object value, string[] parameters)
    {
        if (Handler == null)
        {
            Logger.Warning($"No handler defined for PowerPoint function '{Name}'");
            return $"[Error: Function '{Name}' has no implementation]";
        }

        try
        {
            Logger.Debug($"Executing PowerPoint function '{Name}' with {parameters?.Length ?? 0} parameters");
            var result = Handler(context, value, parameters);
            Logger.Debug($"Function '{Name}' executed successfully");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error executing PowerPoint function '{Name}'", ex);
            return $"[Error in function '{Name}': {ex.Message}]";
        }
    }
}