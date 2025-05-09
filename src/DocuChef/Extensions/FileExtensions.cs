﻿namespace DocuChef.Extensions;

/// <summary>
/// File operation extension methods
/// </summary>
public static class FileExtensions
{
    /// <summary>
    /// Ensures a directory exists for a file path
    /// </summary>
    public static string EnsureDirectoryExists(this string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return filePath;

        string directory = Path.GetDirectoryName(filePath)!;
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return filePath;
    }

    /// <summary>
    /// Gets content type based on file extension
    /// </summary>
    public static string? GetContentType(this string fileExtension)
    {
        if (string.IsNullOrEmpty(fileExtension))
            return null;

        return fileExtension.ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".jpg" => "image/jpeg",
            ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tiff" => "image/tiff",
            ".tif" => "image/tiff",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls" => "application/vnd.ms-excel",
            ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".ppt" => "application/vnd.ms-powerpoint",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".doc" => "application/vnd.ms-word",
            ".pdf" => "application/pdf",
            _ => null
        };
    }

    /// <summary>
    /// Creates a unique file path by adding a counter if file already exists
    /// </summary>
    public static string GetUniquePath(this string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            return filePath;

        string directory = Path.GetDirectoryName(filePath)!;
        string filename = Path.GetFileNameWithoutExtension(filePath);
        string extension = Path.GetExtension(filePath);

        for (int i = 1; i < 1000; i++)
        {
            string newPath = Path.Combine(directory, $"{filename} ({i}){extension}");
            if (!File.Exists(newPath))
                return newPath;
        }

        // If we get here, use a GUID part
        return Path.Combine(directory, $"{filename} ({Guid.NewGuid().ToString("N").Substring(0, 8)}){extension}");
    }

    /// <summary>
    /// Creates a temporary file path with specified extension
    /// </summary>
    public static string GetTempFilePath(this string extension)
    {
        extension = extension.StartsWith(".") ? extension : $".{extension}";
        return Path.Combine(Path.GetTempPath(), $"DocuChef_{Guid.NewGuid().ToString("N")}{extension}");
    }

    /// <summary>
    /// Safely copies a stream to a file, ensuring directory exists
    /// </summary>
    public static void CopyToFile(this Stream source, string destination)
    {
        destination.EnsureDirectoryExists();

        using var fileStream = new FileStream(destination, FileMode.Create, FileAccess.Write);
        source.Position = 0;
        source.CopyTo(fileStream);
    }
}