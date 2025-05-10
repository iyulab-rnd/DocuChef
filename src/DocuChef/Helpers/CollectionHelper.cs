namespace DocuChef.Helpers;

internal static class CollectionHelper
{

    /// <summary>
    /// Helper method to get collection count
    /// </summary>
    public static int GetCollectionCount(object obj)
    {
        if (obj == null)
            return 0;

        if (obj is ICollection collection)
            return collection.Count;

        if (obj is Array array)
            return array.Length;

        // Try to get Count property via reflection
        var countProperty = obj.GetType().GetProperty("Count");
        if (countProperty != null && countProperty.PropertyType == typeof(int) &&
            countProperty.GetGetMethod() != null)
        {
            try
            {
                return (int)countProperty.GetValue(obj);
            }
            catch
            {
                // Fallback to enumerating
            }
        }

        // Try indexer existence to determine if it's a collection
        var indexerProperty = obj.GetType().GetProperty("Item");
        if (indexerProperty != null && indexerProperty.GetIndexParameters().Length > 0)
        {
            try
            {
                // Try to enumerate
                int count = 0;
                var enumerableObj = obj as IEnumerable;
                if (enumerableObj != null)
                {
                    foreach (var _ in enumerableObj)
                        count++;
                    return count;
                }
            }
            catch
            {
                // Fall through to default
            }
        }

        // For any other IEnumerable, count by enumerating
        if (obj is IEnumerable enumerable)
        {
            int count = 0;
            foreach (var _ in enumerable)
                count++;
            return count;
        }

        return 0;
    }
}
