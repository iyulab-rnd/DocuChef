using System.Dynamic;
using System.Reflection;

namespace DocuChef.Utils;

/// <summary>
/// Utility for converting between different data representations
/// </summary>
internal static class DataConverter
{
    /// <summary>
    /// Converts an object to a dictionary of properties and values
    /// </summary>
    public static Dictionary<string, object> ObjectToDictionary(object? obj)
    {
        if (obj == null)
            return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // Already a dictionary
        if (obj is IDictionary<string, object> dictionary)
            return new Dictionary<string, object>(dictionary, StringComparer.OrdinalIgnoreCase);

        // Handle ExpandoObject
        if (obj is ExpandoObject expando)
        {
            IDictionary<string, object?> expandoDict = expando;
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in expandoDict)
            {
                if (kvp.Value != null)
                {
                    result[kvp.Key] = kvp.Value;
                }
                else
                {
                    result[kvp.Key] = string.Empty;
                }
            }
            return result;
        }

        // Handle general dictionary
        if (obj is IDictionary dict)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry entry in dict)
            {
                if (entry.Key != null)
                {
                    result[entry.Key.ToString()!] = entry.Value ?? string.Empty;
                }
            }
            return result;
        }

        // Normal object
        var properties = obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var resultDict = new Dictionary<string, object>(properties.Length, StringComparer.OrdinalIgnoreCase);

        foreach (var prop in properties)
        {
            resultDict[prop.Name] = prop.GetValue(obj) ?? string.Empty;
        }

        return resultDict;
    }

    /// <summary>
    /// Gets the value of a nested property or field from an object
    /// </summary>
    public static object? GetNestedPropertyValue(object? obj, string? propertyPath)
    {
        if (obj == null || string.IsNullOrEmpty(propertyPath))
            return null;

        var parts = propertyPath.Split('.');
        object? current = obj;

        foreach (var part in parts)
        {
            if (current == null)
                return null;

            // Handle dictionary
            if (current is Dictionary<string, object> dict)
            {
                if (dict.TryGetValue(part, out var value))
                {
                    current = value;
                    continue;
                }
                return null;
            }

            // Handle general dictionary
            if (current is IDictionary genDict)
            {
                if (genDict.Contains(part))
                {
                    current = genDict[part];
                    continue;
                }
                return null;
            }

            // Handle regular objects
            var property = current.GetType().GetProperty(part, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (property == null)
                return null;

            current = property.GetValue(current);
        }

        return current;
    }
}