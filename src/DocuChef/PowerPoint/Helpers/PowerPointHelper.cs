using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint.Helpers;

internal static class PowerPointHelper
{
    /// <summary>
    /// 이미지 컨텐츠 타입을 확인하고 적절한 ImagePart를 생성합니다.
    /// </summary>
    public static ImagePart CreateImagePart(SlidePart slidePart, string contentType, string relationshipId)
    {
        switch (contentType)
        {
            case "image/jpeg":
                Logger.Debug("Adding JPEG image part");
                return slidePart.AddImagePart(ImagePartType.Jpeg, relationshipId);
            case "image/png":
                Logger.Debug("Adding PNG image part");
                return slidePart.AddImagePart(ImagePartType.Png, relationshipId);
            case "image/gif":
                Logger.Debug("Adding GIF image part");
                return slidePart.AddImagePart(ImagePartType.Gif, relationshipId);
            case "image/bmp":
                Logger.Debug("Adding BMP image part");
                return slidePart.AddImagePart(ImagePartType.Bmp, relationshipId);
            case "image/tiff":
                Logger.Debug("Adding TIFF image part");
                return slidePart.AddImagePart(ImagePartType.Tiff, relationshipId);
            default:
                Logger.Warning($"Unsupported content type: {contentType}");
                return null;
        }
    }

    /// <summary>
    /// 도형의 외곽선을 복제합니다.
    /// </summary>
    public static A.Outline CloneOutline(A.Outline originalOutline)
    {
        if (originalOutline == null)
            return null;

        Logger.Debug("Cloning outline properties");

        // 깊은 복사를 사용하여 모든 속성과 하위 요소 유지
        return originalOutline.CloneNode(true) as A.Outline;
    }

    /// <summary>
    /// 도형에 대한 기본 외곽선을 생성합니다.
    /// </summary>
    public static A.Outline CreateDefaultOutline(int width = 9525, string colorHex = "000000")
    {
        Logger.Debug($"Creating default outline with width: {width}, color: #{colorHex}");

        var outline = new A.Outline() { Width = width };
        var solidFill = new A.SolidFill();
        var rgbColor = new A.RgbColorModelHex() { Val = colorHex };
        solidFill.AppendChild(rgbColor);
        outline.AppendChild(solidFill);

        return outline;
    }

    /// <summary>
    /// 이미지를 포함하는 Picture 요소를 생성합니다.
    /// </summary>
    public static Picture CreatePicture(
        string relationshipId,
        uint shapeId,
        string shapeName,
        long x,
        long y,
        long width,
        long height,
        bool preserveAspectRatio = true,
        A.Outline outline = null)
    {
        Logger.Debug($"Creating picture: RelID={relationshipId}, ID={shapeId}, Name={shapeName}, " +
                     $"Position=({x}, {y}), Size=({width}, {height})");

        // 새 Picture 요소 생성
        Picture picture = new Picture();

        // NonVisualPictureProperties 설정
        var nvPicProps = new NonVisualPictureProperties(
            new NonVisualDrawingProperties()
            {
                Id = shapeId,
                Name = shapeName
            },
            new NonVisualPictureDrawingProperties(
                new A.PictureLocks() { NoChangeAspect = preserveAspectRatio }
            ),
            new ApplicationNonVisualDrawingProperties()
        );
        picture.AppendChild(nvPicProps);

        // BlipFill 설정
        var blipFill = new BlipFill();
        var blip = new A.Blip() { Embed = relationshipId };
        blipFill.AppendChild(blip);
        blipFill.AppendChild(new A.SourceRectangle());
        var stretch = new A.Stretch();
        stretch.AppendChild(new A.FillRectangle());
        blipFill.AppendChild(stretch);
        picture.AppendChild(blipFill);

        // ShapeProperties 설정
        var shapeProps = new ShapeProperties();
        var transform2D = new A.Transform2D();
        transform2D.Offset = new A.Offset() { X = x, Y = y };
        transform2D.Extents = new A.Extents() { Cx = width, Cy = height };
        shapeProps.AppendChild(transform2D);

        // 도형 형상 설정
        shapeProps.AppendChild(new A.PresetGeometry(
            new A.AdjustValueList()
        )
        { Preset = A.ShapeTypeValues.Rectangle });

        // 외곽선 설정
        if (outline != null)
        {
            shapeProps.AppendChild(outline);
        }

        picture.AppendChild(shapeProps);

        return picture;
    }

    /// <summary>
    /// 슬라이드의 기본 크기를 반환합니다. (기본 16:9 비율)
    /// </summary>
    public static (long Width, long Height) GetDefaultSlideSize()
    {
        // 표준 16:9 슬라이드 크기
        return (9144000, 6858000); // 10인치 x 7.5인치
    }

    /// <summary>
    /// 슬라이드에서 도형 이름으로 도형을 찾습니다.
    /// </summary>
    public static Shape FindShapeByName(SlidePart slidePart, string shapeName)
    {
        if (slidePart?.Slide == null || string.IsNullOrEmpty(shapeName))
            return null;

        foreach (var shape in slidePart.Slide.Descendants<Shape>())
        {
            string currentName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
            if (shapeName == currentName)
                return shape;

            // 대체 방법: Alt Text 확인
            var anvdp = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;
            if (anvdp != null)
            {
                var descAttr = anvdp.GetAttributes()
                    .FirstOrDefault(a => a.LocalName.Equals("descr", StringComparison.OrdinalIgnoreCase));

                if (descAttr.Value == shapeName)
                    return shape;
            }
        }

        return null;
    }
}