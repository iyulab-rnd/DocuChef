using DocuChef.PowerPoint.DollarSignEngine;
using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint;

/// <summary>
/// Main class for PowerPoint template processing
/// </summary>
internal partial class PowerPointProcessor : IExpressionEvaluator
{
    private readonly PresentationDocument _document;
    private readonly PowerPointOptions _options;
    private readonly PowerPointContext _context;
    private readonly ExpressionEvaluator _expressionEvaluator;

    /// <summary>
    /// Initialize PowerPoint processor
    /// </summary>
    public PowerPointProcessor(PresentationDocument document, PowerPointOptions options)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _options = options ?? throw new ArgumentNullException(nameof(options));

        _context = new PowerPointContext { Options = options };
        _expressionEvaluator = new ExpressionEvaluator();

        Logger.Debug("PowerPoint processor initialized");
    }

    /// <summary>
    /// Process PowerPoint template with variables and functions - simplified flow
    /// </summary>
    public void Process(Dictionary<string, object> variables,
                Dictionary<string, Func<object>> globalVariables,
                Dictionary<string, PowerPointFunction> functions)
    {
        // 1. 초기화
        InitializeContext(variables, globalVariables, functions);
        var presentationPart = ValidateDocument();
        var slideIds = GetSlideIds(presentationPart);

        Logger.Info("Phase 1: Analyzing and preparing slides...");

        // 2. 슬라이드 분석 및 복제 (슬라이드 준비 단계)
        foreach (var slideId in slideIds.ToList())  // 순회 중 컬렉션이 변경되므로 복사
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
            AnalyzeAndPrepareSlide(presentationPart, slidePart);
        }

        // 3. 준비된 모든 슬라이드에 바인딩 적용
        Logger.Info("Phase 2: Applying bindings to all slides...");
        var allSlideIds = GetSlideIds(presentationPart);
        foreach (var slideId in allSlideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
            ApplyBindings(slidePart);
        }

        // 최종 저장
        presentationPart.Presentation.Save();
        Logger.Info("PowerPoint template processing completed successfully");
    }

    /// <summary>
    /// Analyze slide and prepare duplicates if needed
    /// </summary>
    private void AnalyzeAndPrepareSlide(PresentationPart presentationPart, SlidePart slidePart)
    {
        string slideId = presentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Analyzing slide {slideId} for array references");

        // 1. 이 슬라이드의 배열 참조 분석
        var arrayReferences = FindArrayReferencesInSlide(slidePart);
        if (!arrayReferences.Any())
        {
            Logger.Debug("No array references found in slide");
            return;
        }

        // 2. 배열별 슬라이드 복제 필요 여부 판단
        foreach (var arrayGroup in arrayReferences.GroupBy(r => r.ArrayName))
        {
            string arrayName = arrayGroup.Key;
            int maxIndex = arrayGroup.Max(r => r.Index);
            int itemsPerSlide = maxIndex + 1;

            Logger.Debug($"Found array '{arrayName}' with max index {maxIndex} in slide");

            // 데이터 배열 크기 확인
            if (!_context.Variables.TryGetValue(arrayName, out var arrayObj) || arrayObj == null)
            {
                Logger.Warning($"Array '{arrayName}' not found in variables");
                continue;
            }

            int totalItems = CollectionHelper.GetCollectionCount(arrayObj);
            Logger.Debug($"Array '{arrayName}' has {totalItems} total items, {itemsPerSlide} items per slide");

            // 복제 필요 여부 판단
            if (totalItems <= itemsPerSlide)
            {
                Logger.Debug($"No duplication needed for array '{arrayName}'");
                continue;  // 복제 불필요
            }

            // 3. 필요한 슬라이드 수 계산 (올림 나눗셈)
            int slidesNeeded = (int)Math.Ceiling((double)totalItems / itemsPerSlide);
            Logger.Info($"Array '{arrayName}' requires {slidesNeeded} slides for {totalItems} items ({itemsPerSlide} items per slide)");

            // 4. 추가 슬라이드 복제 (첫 번째는 이미 있으므로 두 번째부터)
            int baseSlidePosition = SlideHelper.FindSlidePosition(presentationPart, slidePart);

            for (int i = 1; i < slidesNeeded; i++)
            {
                // 배치 시작 인덱스 계산
                int batchStartIndex = i * itemsPerSlide;
                Logger.Debug($"Creating slide {i + 1} for batch starting at index {batchStartIndex}");

                // 슬라이드 복제
                var newSlidePart = SlideHelper.CloneSlide(presentationPart, slidePart);

                // 복제된 슬라이드의 배열 인덱스 업데이트
                UpdateArrayIndices(newSlidePart, arrayName, batchStartIndex);

                // 슬라이드 삽입
                SlideHelper.InsertSlide(presentationPart, newSlidePart, baseSlidePosition + i);
                Logger.Debug($"Inserted duplicated slide at position {baseSlidePosition + i}");

                // 범위를 벗어난 항목 숨김 처리
                HideOutOfRangeItems(newSlidePart, arrayName, batchStartIndex, totalItems);

                // 처리된 슬라이드로 표시
                _context.ProcessedArraySlides.Add(presentationPart.GetIdOfPart(newSlidePart));
            }

            // 원본 슬라이드에도 범위 체크 적용
            HideOutOfRangeItems(slidePart, arrayName, 0, totalItems);
        }

        // 변경 사항 저장
        presentationPart.Presentation.Save();
    }

    /// <summary>
    /// Update array indices in slide with specified offset
    /// </summary>
    private void UpdateArrayIndices(SlidePart slidePart, string arrayName, int offset)
    {
        Logger.Debug($"Updating array indices for '{arrayName}' with offset {offset}");

        int updatedShapeCount = 0;
        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            bool shapeUpdated = false;

            // 모든 텍스트 요소 처리
            var texts = shape.Descendants<A.Text>().ToList();
            foreach (var text in texts)
            {
                if (string.IsNullOrEmpty(text.Text) || !text.Text.Contains(arrayName))
                    continue;

                string original = text.Text;

                // 1. ${ArrayName[n]} 패턴 처리
                var dollarPattern = new Regex($@"\$\{{{arrayName}\[(\d+)\]([^\}}]*)\}}");
                text.Text = dollarPattern.Replace(text.Text, match => {
                    int index = int.Parse(match.Groups[1].Value);
                    string suffix = match.Groups[2].Value;
                    return $"${{{arrayName}[{index + offset}]{suffix}}}";
                });

                // 2. ArrayName[n] 직접 패턴 처리 (함수 인자 등)
                var directPattern = new Regex($@"(?<!\$\{{){arrayName}\[(\d+)\]");
                text.Text = directPattern.Replace(text.Text, match => {
                    int index = int.Parse(match.Groups[1].Value);
                    return $"{arrayName}[{index + offset}]";
                });

                if (text.Text != original)
                {
                    Logger.Debug($"Updated text from '{original}' to '{text.Text}'");
                    shapeUpdated = true;
                }
            }

            if (shapeUpdated)
                updatedShapeCount++;
        }

        Logger.Debug($"Updated array indices in {updatedShapeCount} shapes");
        slidePart.Slide.Save();
    }

    /// <summary>
    /// Hide shapes with out-of-range array indices
    /// </summary>
    private void HideOutOfRangeItems(SlidePart slidePart, string arrayName, int startIndex, int totalItems)
    {
        Logger.Debug($"Checking for out-of-range items in array '{arrayName}': startIndex={startIndex}, totalItems={totalItems}");

        int hiddenShapeCount = 0;
        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (PowerPointShapeHelper.IsShapeHidden(shape))
                continue;

            // 이 도형의 모든 배열 참조 확인
            var references = PowerPointShapeHelper.FindArrayReferences(shape)
                            .Where(r => r.ArrayName == arrayName)
                            .ToList();

            if (!references.Any())
                continue;

            // 범위를 벗어난 참조가 있으면 숨김
            foreach (var reference in references)
            {
                int actualIndex = reference.Index;
                if (actualIndex >= totalItems)
                {
                    Logger.Debug($"Hiding shape '{shape.GetShapeName()}' with reference to {arrayName}[{actualIndex}] (>= {totalItems})");
                    PowerPointShapeHelper.HideShape(shape);
                    hiddenShapeCount++;
                    break;
                }
            }
        }

        Logger.Debug($"Hidden {hiddenShapeCount} shapes with out-of-range references");
        slidePart.Slide.Save();
    }

    /// <summary>
    /// Apply bindings to all visible shapes in slide
    /// </summary>
    private void ApplyBindings(SlidePart slidePart)
    {
        string slideId = _document.PresentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Applying bindings to slide {slideId}");

        // 슬라이드 컨텍스트 설정
        _context.SlidePart = slidePart;
        _context.Slide.Notes = slidePart.GetNotes();

        // 노트에서 디렉티브 처리
        var directives = DirectiveParser.ParseDirectives(_context.Slide.Notes);
        if (directives.Count > 0)
        {
            Logger.Debug($"Processing {directives.Count} directives from slide notes");
            foreach (var directive in directives)
            {
                ProcessShapeDirective(slidePart, directive);
            }
        }

        // 변수 컨텍스트 준비
        var variables = PrepareVariables();
        var textProcessor = new BindingProcessor(this, variables);

        // 모든 도형에 바인딩 적용
        var shapes = slidePart.Slide.Descendants<P.Shape>()
                    .Where(s => !PowerPointShapeHelper.IsShapeHidden(s))
                    .ToList();

        Logger.Debug($"Processing {shapes.Count} visible shapes");

        int processedShapeCount = 0;
        foreach (var shape in shapes)
        {
            bool shapeUpdated = false;

            // 도형 컨텍스트 업데이트
            UpdateShapeContext(shape);

            // 표현식 바인딩
            if (shape.TextBody != null && shape.ContainsExpressions())
            {
                if (textProcessor.ProcessShape(shape))
                {
                    shapeUpdated = true;
                    Logger.Debug($"Processed expressions in shape '{shape.GetShapeName()}'");
                }
            }

            // PowerPoint 함수 처리
            if (ProcessPowerPointFunctions(shape))
            {
                shapeUpdated = true;
                Logger.Debug($"Processed PowerPoint functions in shape '{shape.GetShapeName()}'");
            }

            if (shapeUpdated)
                processedShapeCount++;
        }

        Logger.Debug($"Processed {processedShapeCount} shapes in slide");
        slidePart.Slide.Save();
    }

    /// <summary>
    /// Initialize context with variables and functions
    /// </summary>
    private void InitializeContext(Dictionary<string, object> variables, Dictionary<string, Func<object>> globalVariables, Dictionary<string, PowerPointFunction> functions)
    {
        _context.Variables = variables ?? new Dictionary<string, object>();
        _context.GlobalVariables = globalVariables ?? new Dictionary<string, Func<object>>();
        _context.Functions = functions ?? new Dictionary<string, PowerPointFunction>();
        _context.Variables["_context"] = _context;
    }

    /// <summary>
    /// Validate document structure
    /// </summary>
    private PresentationPart ValidateDocument()
    {
        var presentationPart = _document.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null)
        {
            throw new DocuChefException("Invalid PowerPoint document structure");
        }

        return presentationPart;
    }

    /// <summary>
    /// Get slide IDs from presentation
    /// </summary>
    private List<SlideId> GetSlideIds(PresentationPart presentationPart)
    {
        return presentationPart.Presentation.SlideIdList
            .ChildElements.OfType<SlideId>()
            .ToList();
    }

    /// <summary>
    /// Find all array references in slide
    /// </summary>
    private List<ArrayReference> FindArrayReferencesInSlide(SlidePart slidePart)
    {
        var result = new List<ArrayReference>();

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            var references = PowerPointShapeHelper.FindArrayReferences(shape);
            result.AddRange(references);
        }

        return result;
    }

    /// <summary>
    /// Prepare variables dictionary
    /// </summary>
    internal Dictionary<string, object> PrepareVariables()
    {
        var variables = new Dictionary<string, object>(_context.Variables);

        // Add global variables
        foreach (var globalVar in _context.GlobalVariables)
        {
            variables[globalVar.Key] = globalVar.Value();
        }

        // Add PowerPoint functions
        foreach (var function in _context.Functions)
        {
            variables[$"ppt.{function.Key}"] = function.Value;
        }

        return variables;
    }

    /// <summary>
    /// Evaluate expression with provided variables
    /// </summary>
    public object EvaluateCompleteExpression(string expression, Dictionary<string, object> variables)
    {
        try
        {
            // If already wrapped in ${...}, evaluate directly
            if (expression.StartsWith("${") && expression.EndsWith("}"))
            {
                var result = _expressionEvaluator.Evaluate(expression, variables);
                return result;
            }

            // Otherwise, wrap it for evaluation
            string wrappedExpr = "${" + expression + "}";
            var evalResult = _expressionEvaluator.Evaluate(wrappedExpr, variables);
            return evalResult;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }
}