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
        Logger.Debug($"[IMAGE-DEBUG] Processing image with parameter: '{imagePath}'");

        try
        {
            // Default values from options - 이 부분을 추가합니다
            int width = context.Options?.DefaultImageWidth ?? 300;
            int height = context.Options?.DefaultImageHeight ?? 200;
            bool preserveAspectRatio = context.Options?.PreserveImageAspectRatio ?? true;

            // Check for suspicious array index
            if (imagePath.Contains("Items["))
            {
                var match = System.Text.RegularExpressions.Regex.Match(imagePath, @"Items\[(\d+)\]");
                if (match.Success)
                {
                    int index = int.Parse(match.Groups[1].Value);
                    Logger.Debug($"[IMAGE-DEBUG] Found array index in image path: {index}");

                    // Range check
                    if (index > 30)
                    {
                        Logger.Warning($"[IMAGE-DEBUG] Suspicious high index: {index}, likely an error in index calculation");
                    }

                    // Get the actual array
                    if (context.Variables.TryGetValue("Items", out var arrayObj) && arrayObj != null)
                    {
                        // Try to get the actual count
                        int count = CollectionHelper.GetCollectionCount(arrayObj);

                        if (count > 0 && index >= count)
                        {
                            Logger.Warning($"[IMAGE-DEBUG] Index {index} is out of bounds (count: {count})");

                            // 중요: 범위를 벗어난 경우 null을 반환하여 이미지를 표시하지 않음
                            return "[Out of range]";
                        }
                    }
                }
            }

            // 배열 요소 참조 확인 (예: Items[0].ImageUrl 패턴)
            var arrayIndexMatch = System.Text.RegularExpressions.Regex.Match(
                imagePath,
                @"^(\w+)\[(\d+)\](\.(\w+))+$");

            if (arrayIndexMatch.Success)
            {
                string arrayName = arrayIndexMatch.Groups[1].Value;
                int index = int.Parse(arrayIndexMatch.Groups[2].Value);

                // 전체 속성 경로 추출 (예: .ImageUrl)
                string propertyPath = imagePath.Substring(arrayName.Length + 2 + index.ToString().Length);

                Logger.Debug($"[IMAGE-DEBUG] Detected array reference: array={arrayName}, index={index}, property={propertyPath}");

                // 범위 검사
                if (index > 30)
                {
                    Logger.Warning($"[IMAGE-DEBUG] Index {index} is suspiciously high - might be an error");
                }

                // Items 배열 가져오기
                if (context.Variables.TryGetValue(arrayName, out var arrayObj) && arrayObj != null)
                {
                    // 배열 또는 리스트에서 요소 가져오기
                    object item = null;

                    if (arrayObj is IList list)
                    {
                        // Range check
                        if (index >= 0 && index < list.Count)
                        {
                            item = list[index];
                            Logger.Debug($"[IMAGE-DEBUG] Successfully retrieved item at index {index} from list with {list.Count} items");
                        }
                        else
                        {
                            Logger.Warning($"[IMAGE-DEBUG] Array index out of range: {index}, list count: {list.Count}");
                            // 중요: 범위를 벗어난 경우 null을 반환하여 이미지를 표시하지 않음
                            return "[Out of range]";
                        }
                    }
                    else if (arrayObj is Array array)
                    {
                        if (index >= 0 && index < array.Length)
                        {
                            item = array.GetValue(index);
                            Logger.Debug($"[IMAGE-DEBUG] Successfully retrieved item at index {index} from array with {array.Length} items");
                        }
                        else
                        {
                            Logger.Warning($"[IMAGE-DEBUG] Array index out of range: {index}, array length: {array.Length}");
                            // 중요: 범위를 벗어난 경우 null을 반환하여 이미지를 표시하지 않음
                            return "[Out of range]";
                        }
                    }
                    else if (arrayObj is IEnumerable enumerable)
                    {
                        int count = 0;
                        foreach (var _ in enumerable) count++;

                        int currentIndex = 0;
                        foreach (var enumerableItem in enumerable)
                        {
                            if (currentIndex == index)
                            {
                                item = enumerableItem;
                                Logger.Debug($"[IMAGE-DEBUG] Successfully retrieved item at index {index} from enumerable with {count} items");
                                break;
                            }
                            currentIndex++;
                        }

                        if (item == null)
                        {
                            Logger.Warning($"[IMAGE-DEBUG] Enumerable index out of range: {index}, count: {count}");
                            // 중요: 범위를 벗어난 경우 null을 반환하여 이미지를 표시하지 않음
                            return "[Out of range]";
                        }
                    }

                    if (item != null)
                    {
                        // 속성 경로에서 첫 번째 점 제거
                        if (propertyPath.StartsWith("."))
                            propertyPath = propertyPath.Substring(1);

                        // 리플렉션으로 속성 값 가져오기
                        var props = propertyPath.Split('.');
                        object propValue = item;

                        foreach (var prop in props)
                        {
                            var propInfo = propValue.GetType().GetProperty(prop);
                            if (propInfo == null)
                            {
                                Logger.Warning($"[IMAGE-DEBUG] Property '{prop}' not found on type '{propValue.GetType().Name}'");
                                return $"[Error: Property '{prop}' not found]";
                            }

                            propValue = propInfo.GetValue(propValue);
                            if (propValue == null)
                            {
                                Logger.Warning($"[IMAGE-DEBUG] Property '{prop}' value is null");
                                return $"[Error: Property '{prop}' value is null]";
                            }
                        }

                        // 최종 이미지 경로
                        imagePath = propValue.ToString();
                        Logger.Debug($"[IMAGE-DEBUG] Resolved image path from array item: {imagePath}");
                    }
                    else
                    {
                        Logger.Warning($"[IMAGE-DEBUG] Array item not found: {arrayName}[{index}]");
                        return $"[Error: Array item not found: {arrayName}[{index}]]";
                    }
                }
                else
                {
                    Logger.Warning($"[IMAGE-DEBUG] Array not found: {arrayName}");
                    return $"[Error: Array not found: {arrayName}]";
                }
            }
            // 일반 속성 경로 처리 (예: Object.Property)
            else if (imagePath.Contains("."))
            {
                // Try to resolve as property path
                var resolvedPath = context.ResolveVariable(imagePath);
                if (resolvedPath != null)
                {
                    imagePath = resolvedPath.ToString();
                    Logger.Debug($"[IMAGE-DEBUG] Resolved image path from property path: {imagePath}");
                }
            }
            // 직접 변수 참조 처리
            else if (context.Variables.TryGetValue(imagePath, out var pathObj))
            {
                // Direct variable reference
                imagePath = pathObj?.ToString();
                Logger.Debug($"[IMAGE-DEBUG] Resolved image path from variable: {imagePath}");
            }

            // 여기서 ImageHelper를 사용하여 이미지 처리 (있는 경우)
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
                    // Continue with original path
                }
            }

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