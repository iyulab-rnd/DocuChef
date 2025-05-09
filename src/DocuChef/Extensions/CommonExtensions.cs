namespace DocuChef.Extensions;

/// <summary>
/// Common extension methods used across DocuChef
/// </summary>
public static class CommonExtensions
{
    /// <summary>
    /// Safely gets a property or field value using reflection path notation (e.g. "Customer.Address.City")
    /// </summary>
    public static object ResolvePropertyPath(this object source, string path)
    {
        if (source == null || string.IsNullOrEmpty(path))
            return null;

        // Handle direct property reference
        if (!path.Contains('.'))
        {
            var property = source.GetType().GetProperty(path);
            return property?.GetValue(source);
        }

        // Handle nested properties
        var parts = path.Split('.');
        object current = source;

        foreach (var part in parts)
        {
            if (current == null)
                return null;

            var property = current.GetType().GetProperty(part);
            if (property == null)
                return null;

            current = property.GetValue(current);
        }

        return current;
    }

    /// <summary>
    /// Gets a dictionary of properties and their values from an object
    /// </summary>
    public static Dictionary<string, object> GetProperties(this object source)
    {
        if (source == null)
            return new Dictionary<string, object>();

        var result = new Dictionary<string, object>();
        var type = source.GetType();

        // Get public properties
        foreach (var prop in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
        {
            if (prop.CanRead)
            {
                try
                {
                    var value = prop.GetValue(source);
                    result[prop.Name] = value;
                }
                catch
                {
                    // Skip properties that throw exceptions
                }
            }
        }

        // Get public fields
        foreach (var field in type.GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
        {
            try
            {
                var value = field.GetValue(source);
                result[field.Name] = value;
            }
            catch
            {
                // Skip fields that throw exceptions
            }
        }

        return result;
    }

    /// <summary>
    /// Attempts to convert a dictionary to a strongly typed object
    /// </summary>
    public static T ToObject<T>(this IDictionary<string, object> dictionary) where T : new()
    {
        var result = new T();
        var type = typeof(T);

        foreach (var kvp in dictionary)
        {
            var property = type.GetProperty(kvp.Key);
            if (property != null && property.CanWrite)
            {
                try
                {
                    property.SetValue(result, kvp.Value);
                }
                catch
                {
                    // Skip properties that can't be set
                }
            }
        }

        return result;
    }
}