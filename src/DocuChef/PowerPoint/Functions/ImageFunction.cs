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
        // Expected parameters:
        // 0: Image path or property name that contains image data
        if (parameters == null || parameters.Length == 0)
        {
            Logger.Warning("Image function called without required path parameter");
            return "[Error: Image path required]";
        }

        string imagePath = parameters[0];
        Logger.Debug($"Processing image: {imagePath}");

        try
        {
            // Resolve image path from variables if needed
            if (imagePath.Contains("."))
            {
                // Try to resolve as property path
                var resolvedPath = context.ResolveVariable(imagePath);
                if (resolvedPath != null)
                {
                    imagePath = resolvedPath.ToString();
                    Logger.Debug($"Resolved image path from property path: {imagePath}");
                }
            }
            else if (context.Variables.TryGetValue(imagePath, out var pathObj))
            {
                // Direct variable reference
                imagePath = pathObj?.ToString();
                Logger.Debug($"Resolved image path from variable: {imagePath}");
            }

            // Default values from options
            int width = context.Options?.DefaultImageWidth ?? 300;
            int height = context.Options?.DefaultImageHeight ?? 200;
            bool preserveAspectRatio = context.Options?.PreserveImageAspectRatio ?? true;

            // Parse named parameters according to PPT syntax (width: 300, height: 200)
            for (int i = 1; i < parameters.Length; i++)
            {
                string param = parameters[i];

                // Split by first colon for named parameters
                var colonIndex = param.IndexOf(':');
                if (colonIndex > 0)
                {
                    string paramName = param.Substring(0, colonIndex).Trim();
                    string paramValue = param.Substring(colonIndex + 1).Trim();

                    switch (paramName.ToLowerInvariant())
                    {
                        case "width":
                            if (int.TryParse(paramValue, out int w))
                            {
                                width = w;
                                Logger.Debug($"Custom width parameter: {width}");
                            }
                            break;
                        case "height":
                            if (int.TryParse(paramValue, out int h))
                            {
                                height = h;
                                Logger.Debug($"Custom height parameter: {height}");
                            }
                            break;
                        case "preserveaspectratio":
                            if (bool.TryParse(paramValue, out bool p))
                            {
                                preserveAspectRatio = p;
                                Logger.Debug($"Custom preserveAspectRatio parameter: {preserveAspectRatio}");
                            }
                            break;
                    }
                }
            }

            // Make sure the image file exists
            if (!File.Exists(imagePath))
            {
                Logger.Warning($"Image file not found: {imagePath}");
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
                    return ""; // Return empty string on success to clear the placeholder
                }
                else
                {
                    Logger.Warning("Failed to process image in shape");
                    return "[Image processing failed]";
                }
            }

            Logger.Warning($"Invalid context for image processing - shape: {context.Shape?.ShapeObject != null}, slide part: {context.SlidePart != null}");
            return $"[Image: {imagePath}, {width}x{height}]"; // Debug placeholder
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing image: {ex.Message}", ex);
            if (ex.InnerException != null)
            {
                Logger.Error($"Inner exception: {ex.InnerException.Message}");
            }
            Logger.Error($"Stack trace: {ex.StackTrace}");
            return $"[Error processing image: {ex.Message}]";
        }
    }

    /// <summary>
    /// Process image in shape by replacing the shape with a picture
    /// </summary>
    private static bool ProcessImageInShape(SlidePart slidePart, P.Shape shape, string imagePath, int width, int height, bool preserveAspectRatio)
    {
        try
        {
            // Get existing shape information
            Logger.Debug($"Starting ProcessImageInShape: Path={imagePath}, Shape ID={shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}");

            // Backup existing shape properties
            var nvsp = shape.NonVisualShapeProperties;
            var nvdp = nvsp?.NonVisualDrawingProperties;
            var transform = shape.ShapeProperties?.Transform2D;

            // Position and size information
            uint shapeId = nvdp?.Id?.Value ?? 1000u;
            string shapeName = nvdp?.Name?.Value ?? "Image_Shape";
            long shapeX = transform?.Offset?.X?.Value ?? 1524000;  // Default ~4cm
            long shapeY = transform?.Offset?.Y?.Value ?? 1524000;  // Default ~4cm
            long shapeWidth = transform?.Extents?.Cx?.Value ?? (long)(width * 9525);  // Convert pixels to EMU
            long shapeHeight = transform?.Extents?.Cy?.Value ?? (long)(height * 9525);  // Convert pixels to EMU

            Logger.Debug($"Original shape: ID={shapeId}, Name={shapeName}, X={shapeX}, Y={shapeY}, Width={shapeWidth}, Height={shapeHeight}");

            // Create image part
            string contentType = Path.GetExtension(imagePath).GetContentType();
            if (string.IsNullOrEmpty(contentType))
            {
                Logger.Warning($"Unsupported image format: {Path.GetExtension(imagePath)}");
                return false;
            }

            // Generate relationship ID
            string relationshipId = "R" + Guid.NewGuid().ToString("N").Substring(0, 8);
            Logger.Debug($"Generated relationship ID: {relationshipId}");

            // Create image part
            var imagePart = PowerPointHelper.CreateImagePart(slidePart, contentType, relationshipId);
            if (imagePart == null)
            {
                Logger.Error("Failed to create image part");
                return false;
            }

            // Load image data
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }
            Logger.Debug("Image data fed to part successfully");

            // Get or create outline
            A.Outline outline = null;
            var originalOutline = shape.ShapeProperties?.GetFirstChild<A.Outline>();

            if (originalOutline != null)
            {
                outline = PowerPointHelper.CloneOutline(originalOutline);
                Logger.Debug("Outline cloned from original shape");
            }
            else
            {
                // Optional: Create default outline
                outline = PowerPointHelper.CreateDefaultOutline(19050, "000000"); // 2pt black border
                Logger.Debug("Created default outline");
            }

            // Create picture element
            var picture = PowerPointHelper.CreatePicture(
                relationshipId,
                shapeId + 1, // Ensure unique ID
                shapeName + "_Image",
                shapeX,
                shapeY,
                shapeWidth,
                shapeHeight,
                preserveAspectRatio,
                outline
            );

            // Get parent element
            var parent = shape.Parent;
            if (parent == null)
            {
                Logger.Error("Cannot find parent element for the shape");
                return false;
            }

            // Replace original shape with new picture
            parent.ReplaceChild(picture, shape);
            Logger.Debug("Replaced original shape with new Picture element");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error replacing shape with image: {ex.Message}", ex);
            if (ex.InnerException != null)
            {
                Logger.Error($"Inner exception: {ex.InnerException.Message}");
            }
            Logger.Error($"Stack trace: {ex.StackTrace}");
            return false;
        }
    }
}