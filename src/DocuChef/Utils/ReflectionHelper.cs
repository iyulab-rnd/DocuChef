using System.Reflection;

namespace DocuChef.Utils;

/// <summary>
/// Helper methods for reflection operations
/// </summary>
internal static class ReflectionHelper
{
    /// <summary>
    /// Checks if an object is a collection
    /// </summary>
    public static bool IsCollection(object? obj)
    {
        if (obj == null)
            return false;

        return obj is IEnumerable && obj is not string;
    }

    /// <summary>
    /// Gets the property type for a property path
    /// </summary>
    public static Type? GetPropertyType(Type? type, string? propertyPath)
    {
        if (type == null || string.IsNullOrEmpty(propertyPath))
            return null;

        var parts = propertyPath.Split('.');
        Type? current = type;

        foreach (var part in parts)
        {
            if (current == null)
                return null;

            // Handle dictionary type
            if (typeof(IDictionary<string, object>).IsAssignableFrom(current))
            {
                return typeof(object);
            }

            // Handle general dictionary type
            if (typeof(IDictionary).IsAssignableFrom(current))
            {
                return typeof(object);
            }

            // Handle collection type
            if (typeof(IEnumerable).IsAssignableFrom(current) && current != typeof(string))
            {
                // Try to get element type
                if (current.IsArray)
                {
                    current = current.GetElementType();
                    continue;
                }

                // Try to get generic arguments
                var genericArgs = current.GetGenericArguments();
                if (genericArgs.Length > 0)
                {
                    current = genericArgs[0];
                    continue;
                }

                return typeof(object);
            }

            // Handle regular object type
            var property = current.GetProperty(part, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (property == null)
                return null;

            current = property.PropertyType;
        }

        return current;
    }
}