using DocuChef.PowerPoint.DollarSignEngine;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocuChef.PowerPoint;

/// <summary>
/// PowerPoint 템플릿 처리의 주요 클래스 - 기본 코드
/// </summary>
internal partial class PowerPointProcessor
{
    private readonly PresentationDocument _document;
    private readonly PowerPointOptions _options;
    private readonly PowerPointContext _context;
    private readonly ExpressionEvaluator _expressionEvaluator;
    private readonly TextProcessingHelper _textHelper;

    /// <summary>
    /// PowerPoint 프로세서 초기화
    /// </summary>
    public PowerPointProcessor(PresentationDocument document, PowerPointOptions options)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _options = options ?? throw new ArgumentNullException(nameof(options));

        _context = new PowerPointContext
        {
            Options = options
        };

        // DollarSignEngine 통합을 위한 ExpressionEvaluator 생성
        _expressionEvaluator = new ExpressionEvaluator();

        // 텍스트 처리 도우미 초기화
        _textHelper = new TextProcessingHelper(this, _context);

        Logger.Debug("PowerPoint processor initialized with DollarSignEngine");
    }

    /// <summary>
    /// 변수와 함수로 PowerPoint 템플릿을 처리합니다
    /// </summary>
    public void Process(Dictionary<string, object> variables, Dictionary<string, Func<object>> globalVariables, Dictionary<string, PowerPointFunction> functions)
    {
        // 컨텍스트 설정
        _context.Variables = variables ?? new Dictionary<string, object>();
        _context.GlobalVariables = globalVariables ?? new Dictionary<string, Func<object>>();
        _context.Functions = functions ?? new Dictionary<string, PowerPointFunction>();

        // 특수 컨텍스트 변수 추가
        _context.Variables["_context"] = _context;

        try
        {
            // 프레젠테이션의 모든 슬라이드 가져오기
            var presentationPart = _document.PresentationPart;
            if (presentationPart == null)
            {
                Logger.Error("Invalid PowerPoint document: PresentationPart is missing");
                throw new DocuChefException("Invalid PowerPoint document: PresentationPart is missing");
            }

            var presentation = presentationPart.Presentation;
            if (presentation == null)
            {
                Logger.Error("Invalid PowerPoint document: Presentation is missing");
                throw new DocuChefException("Invalid PowerPoint document: Presentation is missing");
            }

            var slideIdList = presentation.SlideIdList;
            if (slideIdList == null)
            {
                Logger.Error("Invalid PowerPoint document: SlideIdList is missing");
                throw new DocuChefException("Invalid PowerPoint document: SlideIdList is missing");
            }

            // 각 슬라이드 처리
            var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
            Logger.Info($"Processing {slideIds.Count} slides");

            for (int i = 0; i < slideIds.Count; i++)
            {
                ProcessSlide(presentationPart, slideIds[i], i);

                // 이 슬라이드에 대한 모든 변경사항이 저장되도록 합니다
                try
                {
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideIds[i].RelationshipId);
                    if (slidePart != null && slidePart.Slide != null)
                    {
                        slidePart.Slide.Save();
                        Logger.Debug($"Slide {i} saved after processing");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error saving slide {i}: {ex.Message}", ex);
                }
            }

            // 모든 슬라이드 처리 후 프레젠테이션 저장
            try
            {
                _document.PresentationPart.Presentation.Save();
                Logger.Info("Presentation saved after processing all slides");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error saving presentation: {ex.Message}", ex);
            }

            Logger.Info("PowerPoint template processing completed");
        }
        catch (Exception ex)
        {
            Logger.Error("Error processing PowerPoint template", ex);
            throw new TemplateProcessingException($"Error processing PowerPoint template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 단일 슬라이드 처리
    /// </summary>
    private void ProcessSlide(PresentationPart presentationPart, SlideId slideId, int slideIndex)
    {
        Logger.Debug($"Processing slide {slideIndex} with ID {slideId.RelationshipId}");

        // 슬라이드 컨텍스트 업데이트
        _context.Slide.Index = slideIndex;
        _context.Slide.Id = slideId.RelationshipId;

        // 슬라이드 파트 가져오기
        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
        if (slidePart == null || slidePart.Slide == null)
        {
            Logger.Warning($"Slide part not found for ID {slideId.RelationshipId}");
            return;
        }

        // 컨텍스트에 SlidePart 저장
        _context.SlidePart = slidePart;

        // 슬라이드 노트 가져오기
        string slideNotes = slidePart.GetNotes();
        _context.Slide.Notes = slideNotes;

        Logger.Debug($"Slide notes: {slideNotes}");

        // 향상된 DirectiveParser를 사용하여 슬라이드 노트에서 지시문 파싱
        var directives = DirectiveParser.ParseDirectives(slideNotes);

        // foreach 지시문 먼저 확인 - 이것이 슬라이드 복제를 처리합니다
        var foreachDirective = directives.FirstOrDefault(d => d.Name == "foreach");
        if (foreachDirective != null)
        {
            Logger.Debug($"Found foreach directive, processing for slide duplication");
            ProcessSlideDirective(presentationPart, slidePart, foreachDirective);
            return; // 지시문이 이 슬라이드를 처리하므로 추가 처리는 건너뜁니다
        }

        // 다른 지시문 처리 (예: #if)
        foreach (var directive in directives.Where(d => d.Name != "foreach"))
        {
            Logger.Debug($"Processing directive: {directive.Name}");
            ProcessShapeDirective(slidePart, directive);
        }

        // DollarSignEngine을 사용하여 텍스트 교체 처리
        Logger.Debug($"Processing text replacements with DollarSignEngine on slide {slideIndex}");
        ProcessTextReplacements(slidePart);

        // 처리 후 슬라이드 저장
        try
        {
            slidePart.Slide.Save();
            Logger.Debug($"Slide {slideIndex} saved after processing");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error saving slide {slideIndex}: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 컨텍스트 변수와 전역 변수를 결합한 변수 사전 준비
    /// </summary>
    internal Dictionary<string, object> PrepareVariables()
    {
        var variables = new Dictionary<string, object>(_context.Variables);

        // 전역 변수 추가
        foreach (var globalVar in _context.GlobalVariables)
        {
            variables[globalVar.Key] = globalVar.Value();
        }

        // PowerPoint 함수 추가
        foreach (var function in _context.Functions)
        {
            variables[$"ppt.{function.Key}"] = function.Value;
        }

        return variables;
    }

    /// <summary>
    /// DollarSignEngine을 사용하여 표현식 평가
    /// </summary>
    internal object EvaluateCompleteExpression(string expression)
    {
        // 변수 사전 준비
        var variables = PrepareVariables();
        return EvaluateCompleteExpression(expression, variables);
    }

    /// <summary>
    /// 제공된 변수로 표현식 평가
    /// </summary>
    internal object EvaluateCompleteExpression(string expression, Dictionary<string, object> variables)
    {
        // DollarSignEngine 어댑터를 사용하여 평가
        try
        {
            return _expressionEvaluator.Evaluate(expression, variables);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// 컨텍스트에서 변수 값 해결
    /// </summary>
    private object ResolveVariableValue(string name)
    {
        // 직접 변수 확인
        if (_context.Variables.TryGetValue(name, out var value))
            return value;

        // 전역 변수 확인
        if (_context.GlobalVariables.TryGetValue(name, out var factory))
            return factory();

        // 속성 경로 확인
        if (name.Contains('.'))
        {
            var parts = name.Split('.');
            if (_context.Variables.TryGetValue(parts[0], out var obj))
            {
                for (int i = 1; i < parts.Length && obj != null; i++)
                {
                    var property = obj.GetType().GetProperty(parts[i]);
                    obj = property?.GetValue(obj);
                }
                return obj;
            }
        }

        return null;
    }
}

/// <summary>
/// 텍스트 처리를 위한 헬퍼 클래스
/// </summary>
internal class TextProcessingHelper
{
    private readonly PowerPointProcessor _processor;
    private readonly PowerPointContext _context;

    /// <summary>
    /// 텍스트 처리 헬퍼 초기화
    /// </summary>
    public TextProcessingHelper(PowerPointProcessor processor, PowerPointContext context)
    {
        _processor = processor;
        _context = context;
    }

    /// <summary>
    /// 슬라이드의 텍스트 교체 처리 - 개선된 알고리즘
    /// </summary>
    public void ProcessTextReplacements(SlidePart slidePart)
    {
        // 슬라이드의 모든 도형 요소 가져오기
        var shapes = slidePart.Slide.Descendants<Shape>().ToList();
        Logger.Debug($"Processing text replacements in {shapes.Count} shapes");

        bool hasTextChanges = false;

        foreach (var shape in shapes)
        {
            // 도형 이름 가져오기
            string shapeName = shape.GetShapeName();
            Logger.Debug($"Processing shape: {shapeName ?? "(unnamed)"}");

            // 도형 컨텍스트 업데이트
            UpdateShapeContext(shape);

            // 이 도형의 전체 텍스트 콘텐츠 재구성
            string completeText = ReconstructCompleteText(shape);
            if (string.IsNullOrEmpty(completeText))
                continue;

            Logger.Debug($"Complete text in shape: '{completeText}'");

            // 표현식이나 함수가 포함되어 있는지 확인
            if (ContainsExpressions(completeText))
            {
                try
                {
                    // 텍스트 단위로 처리하는 대신 전체 도형 텍스트를 한 번에 처리
                    string processedText = ProcessCompleteText(completeText);
                    if (processedText != completeText)
                    {
                        Logger.Debug($"Replacing complete text: '{completeText}' -> '{processedText}'");

                        // 처리된 텍스트로 도형 내용 업데이트
                        UpdateShapeText(shape, processedText);
                        hasTextChanges = true;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error processing shape text: {ex.Message}", ex);
                    UpdateShapeText(shape, $"[Error: {ex.Message}]");
                    hasTextChanges = true;
                }
            }
        }

        // 모든 텍스트 교체 후 변경사항이 있으면 슬라이드 저장
        if (hasTextChanges)
        {
            try
            {
                slidePart.Slide.Save();
                Logger.Debug("Slide saved after text replacements");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error saving slide after text replacements: {ex.Message}", ex);
            }
        }
    }

    /// <summary>
    /// 도형의 완전한 텍스트 재구성
    /// </summary>
    private string ReconstructCompleteText(Shape shape)
    {
        // 모든 텍스트 실행 가져오기
        var paragraphs = shape.Descendants<A.Paragraph>().ToList();
        if (paragraphs.Count == 0)
            return string.Empty;

        StringBuilder sb = new StringBuilder();

        foreach (var paragraph in paragraphs)
        {
            StringBuilder paragraphText = new StringBuilder();

            // 단락 내 모든 텍스트 실행 병합
            foreach (var run in paragraph.Descendants<A.Run>())
            {
                var text = run.Descendants<A.Text>().FirstOrDefault();
                if (text != null && !string.IsNullOrEmpty(text.Text))
                {
                    paragraphText.Append(text.Text);
                }
            }

            // 단락 사이에 줄바꿈 추가
            if (sb.Length > 0 && paragraphText.Length > 0)
            {
                sb.AppendLine();
            }

            sb.Append(paragraphText);
        }

        return sb.ToString();
    }

    /// <summary>
    /// 도형의 텍스트를 서식을 보존하면서 업데이트
    /// </summary>
    private void UpdateShapeText(Shape shape, string newText)
    {
        if (shape.TextBody == null)
            return;

        try
        {
            // P.TextBody를 사용 (DocumentFormat.OpenXml.Presentation)
            var textBody = shape.TextBody;

            // 기존 단락 및 서식 정보 가져오기
            var existingParagraphs = textBody.Elements<A.Paragraph>().ToList();
            if (existingParagraphs.Count == 0)
            {
                // 기존 단락이 없으면 기본 스타일로 새 단락 생성
                CreateNewParagraphsFromText(shape, newText);
                return;
            }

            // 첫 번째 단락의 스타일 정보 저장 (정렬, 들여쓰기 등)
            var firstParaProps = existingParagraphs[0].ParagraphProperties?.CloneNode(true) as A.ParagraphProperties;

            // 첫 번째 실행 요소의 스타일 정보 저장 (폰트, 크기, 굵게, 기울임 등)
            var firstRun = existingParagraphs[0].Descendants<A.Run>().FirstOrDefault();
            var firstRunProps = firstRun?.RunProperties?.CloneNode(true) as A.RunProperties;

            // 기존 단락 모두 제거
            foreach (var para in existingParagraphs)
            {
                para.Remove();
            }

            // 여러 줄로 텍스트 분할
            string[] lines = newText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0)
                lines = new[] { newText }; // 줄바꿈이 없으면 전체 텍스트를 한 줄로 처리

            // 각 줄에 대해 스타일이 적용된 새 단락 생성
            foreach (var line in lines)
            {
                var paragraph = new A.Paragraph();

                // 단락 속성 적용 (정렬 등)
                if (firstParaProps != null)
                {
                    paragraph.ParagraphProperties = firstParaProps.CloneNode(true) as A.ParagraphProperties;
                }

                var run = new A.Run();

                // 실행 속성 적용 (폰트, 크기, 굵게, 기울임 등)
                if (firstRunProps != null)
                {
                    run.RunProperties = firstRunProps.CloneNode(true) as A.RunProperties;
                }

                var text = new A.Text(line);

                run.AppendChild(text);
                paragraph.AppendChild(run);
                textBody.AppendChild(paragraph);
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error updating shape text with styles: {ex.Message}", ex);

            // 오류 발생 시 기본 방식으로 대체
            try
            {
                CreateNewParagraphsFromText(shape, newText);
            }
            catch (Exception fallbackEx)
            {
                Logger.Error($"Fallback text update also failed: {fallbackEx.Message}", fallbackEx);
            }
        }
    }

    /// <summary>
    /// 새 텍스트로 기본 스타일의 단락 생성
    /// </summary>
    private void CreateNewParagraphsFromText(Shape shape, string text)
    {
        if (shape.TextBody == null)
            return;

        var textBody = shape.TextBody;

        // 기존 단락 모두 제거
        var existingParagraphs = textBody.Elements<A.Paragraph>().ToList();
        foreach (var para in existingParagraphs)
        {
            para.Remove();
        }

        // 여러 줄로 텍스트 분할
        string[] lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0)
            lines = new[] { text }; // 줄바꿈이 없으면 전체 텍스트를 한 줄로 처리

        // 각 줄에 대해 새 단락 생성
        foreach (string line in lines)
        {
            var paragraph = new A.Paragraph();
            var run = new A.Run();
            var textElement = new A.Text(line);

            run.AppendChild(textElement);
            paragraph.AppendChild(run);
            textBody.AppendChild(paragraph);
        }
    }

    /// <summary>
    /// 도형 컨텍스트 업데이트
    /// </summary>
    private void UpdateShapeContext(Shape shape)
    {
        string name = shape.GetShapeName();
        string text = shape.GetText();

        _context.Shape = new ShapeContext
        {
            Name = name,
            Id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
            Text = text,
            Type = GetShapeType(shape),
            ShapeObject = shape
        };
    }

    /// <summary>
    /// 도형 유형 가져오기
    /// </summary>
    private string GetShapeType(Shape shape)
    {
        // 도형 속성 확인
        if (shape.ShapeProperties != null)
        {
            // PresetGeometry 찾기
            var presetGeometry = shape.ShapeProperties.ChildElements
                                    .OfType<A.PresetGeometry>()
                                    .FirstOrDefault();

            if (presetGeometry?.Preset != null)
            {
                return presetGeometry.Preset.Value.ToString();
            }
        }

        // TextBody 확인
        if (shape.TextBody != null)
        {
            return "TextBox";
        }

        return "Shape";
    }

    /// <summary>
    /// 텍스트에 표현식이나 함수가 포함되어 있는지 확인
    /// </summary>
    private bool ContainsExpressions(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        // ${...} 패턴 체크
        if (text.Contains("${"))
            return true;

        // 배열 인덱스 패턴 체크 (예: item[0].Id)
        return Regex.IsMatch(text, @"\b\w+\[\d+\](\.[\w]+)*\b");
    }

    /// <summary>
    /// 완전한 텍스트 처리 - 모든 표현식 평가
    /// </summary>
    private string ProcessCompleteText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        try
        {
            // 준비된 변수 딕셔너리 가져오기
            var variables = _processor.PrepareVariables();

            // 컨텍스트 변수 추가
            variables["_context"] = _context;

            // 1. ${...} 형식의 표현식 처리
            var dollarExprPattern = @"\${([^{}]+)}";
            text = ProcessExpressionPattern(text, dollarExprPattern, variables);

            // 2. 포맷 지정자가 있는 표현식 특별 처리: ${item[0].Price:C0}
            var formatExprPattern = @"\${([^:{}]+):([^{}]+)}";
            text = ProcessFormattedExpressions(text, formatExprPattern, variables);

            // 3. 배열 인덱스 패턴 처리: item[0].Id
            var arrayIndexPattern = @"\b(\w+)\[(\d+)\](\.(\w+))?";
            text = ProcessArrayIndexExpressions(text, arrayIndexPattern, variables);

            return text;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing text: {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// 패턴 매칭을 사용하여 텍스트의 모든 표현식 처리
    /// </summary>
    private string ProcessExpressionPattern(string text, string pattern, Dictionary<string, object> variables)
    {
        return Regex.Replace(text, pattern, match => {
            string expr = match.Value;
            try
            {
                var result = _processor.EvaluateCompleteExpression(expr, variables);
                return result?.ToString() ?? "";
            }
            catch
            {
                return expr; // 오류 시 원본 유지
            }
        });
    }

    /// <summary>
    /// 포맷 지정자가 있는 표현식 처리 (예: ${item[0].Price:C0})
    /// </summary>
    private string ProcessFormattedExpressions(string text, string pattern, Dictionary<string, object> variables)
    {
        return Regex.Replace(text, pattern, match => {
            try
            {
                string valueExpr = match.Groups[1].Value; // item[0].Price
                string format = match.Groups[2].Value; // C0

                // 값 표현식 평가
                string dollarExpr = $"${{{valueExpr}}}";
                var value = _processor.EvaluateCompleteExpression(dollarExpr, variables);

                if (value == null)
                    return "";

                // 값 유형에 따라 포맷 적용
                if (value is IFormattable formattable)
                {
                    try
                    {
                        return formattable.ToString(format, System.Globalization.CultureInfo.CurrentCulture);
                    }
                    catch
                    {
                        return value.ToString();
                    }
                }

                return value.ToString();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error formatting expression: {ex.Message}");
                return match.Value; // 오류 시 원본 유지
            }
        });
    }

    /// <summary>
    /// 배열 인덱스 표현식 처리 (예: item[0].Id)
    /// </summary>
    private string ProcessArrayIndexExpressions(string text, string pattern, Dictionary<string, object> variables)
    {
        return Regex.Replace(text, pattern, match => {
            try
            {
                string arrayName = match.Groups[1].Value; // item
                string indexStr = match.Groups[2].Value; // 0
                string propPath = match.Groups[4].Success ? match.Groups[4].Value : null; // Id or null

                // 달러 기호 표현식으로 변환
                string dollarExpr;
                if (propPath != null)
                    dollarExpr = $"${{{arrayName}[{indexStr}].{propPath}}}";
                else
                    dollarExpr = $"${{{arrayName}[{indexStr}]}}";

                // 표현식 평가
                var result = _processor.EvaluateCompleteExpression(dollarExpr, variables);
                return result?.ToString() ?? "";
            }
            catch
            {
                return match.Value; // 오류 시 원본 유지
            }
        });
    }
}