using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Functions;

/// <summary>
/// Image-related functions for PowerPoint processing according to PPT syntax guidelines
/// </summary>
internal static class ImageFunction
{
    /// <summary>
    /// Creates a PowerPoint function for image handling
    /// </summary>
    public static PowerPointFunction Create()
    {
        return new PowerPointFunction
        {
            Name = "Image",
            Description = "Inserts an image into a PowerPoint shape according to ppt.Image syntax",
            Handler = ProcessImageFunction
        };
    }

    /// <summary>
    /// Process image function: ppt.Image("imageProperty", width: 300, height: 200, preserveAspectRatio: true)
    /// </summary>
    private static object ProcessImageFunction(PowerPointContext context, object value, string[] parameters)
    {
        if (parameters == null || parameters.Length == 0)
        {
            Logger.Warning("Image function called without required path parameter");
            return "[Error: Image path required]";
        }

        string imagePath = parameters[0];
        Logger.Debug($"[IMAGE-DEBUG] Processing image with parameter: '{imagePath}'");

        try
        {
            int width = context.Options?.DefaultImageWidth ?? 300;
            int height = context.Options?.DefaultImageHeight ?? 200;
            bool preserveAspectRatio = context.Options?.PreserveImageAspectRatio ?? true;

            // Resolve array references
            if (imagePath.Contains('[') && imagePath.Contains(']'))
            {
                imagePath = ResolveArrayIndexedPath(context, imagePath);
            }
            // Resolve property paths
            else if (imagePath.Contains('.'))
            {
                imagePath = ResolvePropertyPath(context, imagePath);
            }
            // Direct variable reference
            else if (context.Variables.TryGetValue(imagePath, out var pathObj))
            {
                imagePath = pathObj?.ToString();
            }

            // Use ImageHelper if available
            if (ClosedXML.Report.XLCustom.Functions.ImageHelper.GetImageFromPathOrUrl != null)
            {
                try
                {
                    string resolvedImagePath = ClosedXML.Report.XLCustom.Functions.ImageHelper.GetImageFromPathOrUrl(imagePath);
                    if (!string.IsNullOrEmpty(resolvedImagePath))
                    {
                        Logger.Debug($"[IMAGE-DEBUG] Image resolved with ImageHelper: {resolvedImagePath}");
                        imagePath = resolvedImagePath;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"[IMAGE-DEBUG] Error using ImageHelper: {ex.Message}");
                }
            }

            // Parse named parameters
            ParseImageParameters(parameters, ref width, ref height, ref preserveAspectRatio);

            // Validate image file exists
            if (!File.Exists(imagePath))
            {
                Logger.Warning($"[IMAGE-DEBUG] Image file not found: {imagePath}");
                return $"[Error: Image file not found: {imagePath}]";
            }

            var fileInfo = new FileInfo(imagePath);
            Logger.Debug($"Image file exists: {imagePath}, size: {fileInfo.Length} bytes");

            // Process the image in the shape
            if (context.Shape?.ShapeObject != null && context.SlidePart != null)
            {
                Logger.Debug($"Processing image in shape: File={imagePath}, Width={width}, Height={height}, PreserveAspectRatio={preserveAspectRatio}");
                Logger.Debug($"Shape ID: {context.Shape.ShapeObject.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}");
                Logger.Debug($"Shape Name: {context.Shape.Name}");

                bool success = ProcessImageInShape(context.SlidePart, context.Shape.ShapeObject, imagePath, width, height, preserveAspectRatio);

                if (success)
                {
                    Logger.Info("Successfully processed image in shape");
                    return "";
                }
                else
                {
                    Logger.Warning("Failed to process image in shape");
                    return "[Image processing failed]";
                }
            }

            Logger.Warning($"Invalid context for image processing - shape: {context.Shape?.ShapeObject != null}, slide part: {context.SlidePart != null}");
            return $"[Image: {imagePath}, {width}x{height}]";
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing image: {ex.Message}", ex);
            return $"[Error processing image: {ex.Message}]";
        }
    }

    /// <summary>
    /// Resolve array indexed path (e.g., Items[0].ImageUrl)
    /// </summary>
    private static string ResolveArrayIndexedPath(PowerPointContext context, string path)
    {
        var match = System.Text.RegularExpressions.Regex.Match(path, @"^(\w+)\[(\d+)\](\.(.+))?$");
        if (!match.Success)
            return path;

        string arrayName = match.Groups[1].Value;
        int index = int.Parse(match.Groups[2].Value);
        var propertyPath = match.Groups[4].Success ? match.Groups[4].Value : null;

        Logger.Debug($"[IMAGE-DEBUG] Detected array reference: array={arrayName}, index={index}, property={propertyPath}");

        if (!context.Variables.TryGetValue(arrayName, out var arrayObj) || arrayObj == null)
        {
            Logger.Warning($"[IMAGE-DEBUG] Array not found: {arrayName}");
            return $"[Error: Array not found: {arrayName}]";
        }

        object item = CollectionHelper.GetItemAtIndex(arrayObj, index);
        if (item == null)
        {
            Logger.Warning($"[IMAGE-DEBUG] Array item not found: {arrayName}[{index}]");
            return $"[Error: Array item not found: {arrayName}[{index}]]";
        }

        Logger.Debug($"[IMAGE-DEBUG] Successfully retrieved item at index {index}");

        if (string.IsNullOrEmpty(propertyPath))
            return item.ToString();

        object propValue = ResolveNestedProperty(item, propertyPath);
        if (propValue == null)
        {
            Logger.Warning($"[IMAGE-DEBUG] Property '{propertyPath}' not found or null");
            return $"[Error: Property '{propertyPath}' not found]";
        }

        return propValue.ToString();
    }

    /// <summary>
    /// Resolve property path
    /// </summary>
    private static string ResolvePropertyPath(PowerPointContext context, string path)
    {
        var resolvedPath = context.ResolveVariable(path);
        if (resolvedPath != null)
        {
            Logger.Debug($"[IMAGE-DEBUG] Resolved image path from property path: {resolvedPath}");
            return resolvedPath.ToString();
        }
        return path;
    }

    /// <summary>
    /// Resolve nested property from object
    /// </summary>
    private static object? ResolveNestedProperty(object obj, string propertyPath)
    {
        var props = propertyPath.Split('.');
        object? current = obj;

        foreach (var prop in props)
        {
            if (current == null)
                return null;

            var property = current.GetType().GetProperty(prop);
            if (property == null)
                return null;

            current = property.GetValue(current);
        }

        return current;
    }

    /// <summary>
    /// Parse image parameters
    /// </summary>
    private static void ParseImageParameters(string[] parameters, ref int width, ref int height, ref bool preserveAspectRatio)
    {
        for (int i = 1; i < parameters.Length; i++)
        {
            string param = parameters[i];
            int colonIndex = param.IndexOf(':');
            if (colonIndex <= 0)
                continue;

            string paramName = param.Substring(0, colonIndex).Trim();
            string paramValue = param.Substring(colonIndex + 1).Trim();

            switch (paramName.ToLowerInvariant())
            {
                case "width":
                    if (int.TryParse(paramValue, out int w))
                        width = w;
                    break;
                case "height":
                    if (int.TryParse(paramValue, out int h))
                        height = h;
                    break;
                case "preserveaspectratio":
                    if (bool.TryParse(paramValue, out bool p))
                        preserveAspectRatio = p;
                    break;
            }
        }
    }

    /// <summary>
    /// Process image in shape by replacing the shape with a picture
    /// </summary>
    private static bool ProcessImageInShape(SlidePart slidePart, P.Shape shape, string imagePath, int width, int height, bool preserveAspectRatio)
    {
        try
        {
            Logger.Debug($"Starting ProcessImageInShape: Path={imagePath}, Shape ID={shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}");

            var nvsp = shape.NonVisualShapeProperties;
            var nvdp = nvsp?.NonVisualDrawingProperties;
            var transform = shape.ShapeProperties?.Transform2D;

            uint shapeId = nvdp?.Id?.Value ?? 1000u;
            string shapeName = nvdp?.Name?.Value ?? "Image_Shape";
            long shapeX = transform?.Offset?.X?.Value ?? 1524000;
            long shapeY = transform?.Offset?.Y?.Value ?? 1524000;
            long shapeWidth = transform?.Extents?.Cx?.Value ?? (long)(width * 9525);
            long shapeHeight = transform?.Extents?.Cy?.Value ?? (long)(height * 9525);

            Logger.Debug($"Original shape: ID={shapeId}, Name={shapeName}, X={shapeX}, Y={shapeY}, Width={shapeWidth}, Height={shapeHeight}");

            string contentType = Path.GetExtension(imagePath).GetContentType();
            if (string.IsNullOrEmpty(contentType))
            {
                Logger.Warning($"Unsupported image format: {Path.GetExtension(imagePath)}");
                return false;
            }

            string relationshipId = GenerateUniqueRelationshipId(slidePart);
            Logger.Debug($"Generated relationship ID: {relationshipId}");

            var imagePart = PowerPointHelper.CreateImagePart(slidePart, contentType, relationshipId);
            if (imagePart == null)
            {
                Logger.Error("Failed to create image part");
                return false;
            }

            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }
            Logger.Debug("Image data fed to part successfully");

            A.Outline outline = null;
            var originalOutline = shape.ShapeProperties?.GetFirstChild<A.Outline>();

            if (originalOutline != null)
            {
                outline = PowerPointHelper.CloneOutline(originalOutline);
                Logger.Debug("Outline cloned from original shape");
            }
            else
            {
                outline = PowerPointHelper.CreateDefaultOutline(9525, "808080");
                Logger.Debug("Created default outline");
            }

            var picture = PowerPointHelper.CreatePicture(
                relationshipId,
                GetNextShapeId(slidePart),
                shapeName + "_Image",
                shapeX,
                shapeY,
                shapeWidth,
                shapeHeight,
                preserveAspectRatio,
                outline
            );

            var parent = shape.Parent;
            if (parent == null)
            {
                Logger.Error("Cannot find parent element for the shape");
                return false;
            }

            parent.InsertAfter(picture, shape);
            parent.RemoveChild(shape);

            Logger.Debug("Replaced original shape with new Picture element");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error replacing shape with image: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Generate unique relationship ID for slide part
    /// </summary>
    private static string GenerateUniqueRelationshipId(SlidePart slidePart)
    {
        string baseId = "rImage";
        int counter = 1;

        var existingIds = slidePart.Parts.Select(p => p.RelationshipId).ToHashSet();

        string relationshipId;
        do
        {
            relationshipId = $"{baseId}{counter++}";
        } while (existingIds.Contains(relationshipId));

        return relationshipId;
    }

    /// <summary>
    /// Get next available shape ID
    /// </summary>
    private static uint GetNextShapeId(SlidePart slidePart)
    {
        uint maxId = 0;

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            var id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id.HasValue && id.Value > maxId)
                maxId = id.Value;
        }

        foreach (var pic in slidePart.Slide.Descendants<Picture>())
        {
            var id = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id.HasValue && id.Value > maxId)
                maxId = id.Value;
        }

        return maxId + 1;
    }
}