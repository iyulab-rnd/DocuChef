namespace DocuChef.Utils;

/// <summary>
/// Provides utility methods for DocuChef
/// </summary>
internal static class DocuChefUtils
{
    /// <summary>
    /// Gets the file extension from a path
    /// </summary>
    public static string GetFileExtension(string path)
    {
        if (string.IsNullOrEmpty(path))
            return string.Empty;

        return Path.GetExtension(path).ToLowerInvariant();
    }

    /// <summary>
    /// Checks if a file exists and throws a FileNotFoundException if it doesn't
    /// </summary>
    public static void EnsureFileExists(string filePath, string paramName = null)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(paramName ?? nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}", filePath);
    }

    /// <summary>
    /// Creates directory for a file path if it doesn't exist
    /// </summary>
    public static void EnsureDirectoryExists(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return;

        string directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    /// <summary>
    /// Converts a pascal case string to a space-separated string
    /// </summary>
    public static string PascalCaseToSpaced(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        StringBuilder result = new StringBuilder();
        result.Append(text[0]);

        for (int i = 1; i < text.Length; i++)
        {
            if (char.IsUpper(text[i]))
                result.Append(' ');

            result.Append(text[i]);
        }

        return result.ToString();
    }

    /// <summary>
    /// Gets a unique filename in the specified directory
    /// </summary>
    public static string GetUniqueFilename(string directory, string filename)
    {
        if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(filename))
            return filename;

        string filePath = Path.Combine(directory, filename);
        if (!File.Exists(filePath))
            return filename;

        string name = Path.GetFileNameWithoutExtension(filename);
        string extension = Path.GetExtension(filename);
        int counter = 1;

        do
        {
            string newFilename = $"{name} ({counter}){extension}";
            filePath = Path.Combine(directory, newFilename);
            counter++;

            if (!File.Exists(filePath))
                return newFilename;
        }
        while (counter < 1000); // Prevent infinite loop

        return $"{name} ({Guid.NewGuid().ToString().Substring(0, 8)}){extension}";
    }
}