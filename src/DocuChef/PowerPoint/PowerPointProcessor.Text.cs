using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint;

/// <summary>
/// PowerPointProcessor 부분 클래스 - 텍스트 처리 메서드
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// 슬라이드의 텍스트 교체 처리
    /// </summary>
    private void ProcessTextReplacements(SlidePart slidePart)
    {
        // 텍스트 처리 헬퍼를 통해 모든 도형의 텍스트 처리
        _textHelper.ProcessTextReplacements(slidePart);
    }

    /// <summary>
    /// PowerPoint 도형에서 특수 함수 처리 (이미지, 차트, 표 등)
    /// </summary>
    private bool ProcessPowerPointFunction(P.Shape shape, A.Text textRun)
    {
        string text = textRun.Text;
        Logger.Debug($"Processing PowerPoint function: {text}");
        bool textModified = false;

        try
        {
            // 변수 딕셔너리 준비
            var variables = PrepareVariables();

            // 텍스트에서 모든 ppt. 함수 표현식 추출
            var matches = Regex.Matches(text, @"\${ppt\.(\w+)\(([^)]*)\)}");

            if (matches.Count > 0)
            {
                // 전체 텍스트가 단일 함수 호출인 경우
                if (matches.Count == 1 && matches[0].Value == text)
                {
                    string functionName = matches[0].Groups[1].Value;
                    string parametersString = matches[0].Groups[2].Value;

                    Logger.Debug($"Function: {functionName}, Parameters: {parametersString}");

                    // 함수가 존재하면 실행
                    if (_context.Functions.TryGetValue(functionName, out var function))
                    {
                        // 도형 컨텍스트 업데이트
                        _context.Shape.ShapeObject = shape;

                        // 매개변수 파싱
                        var parameters = ParseFunctionParameters(parametersString);

                        // 함수 핸들러 호출
                        Logger.Debug($"Executing function {functionName} with parameters: {string.Join(", ", parameters)}");
                        var result = function.Execute(_context, null, parameters);

                        // 함수 결과 처리
                        if (result is string resultText)
                        {
                            if (string.IsNullOrEmpty(resultText))
                            {
                                // 성공 케이스 (예: 이미지가 성공적으로 처리됨)
                                textRun.Text = "";
                                Logger.Debug($"Function {functionName} executed successfully with empty result");
                            }
                            else
                            {
                                // 결과 텍스트 또는 오류 메시지
                                textRun.Text = resultText;
                                Logger.Debug($"Function {functionName} result: {resultText}");
                            }
                            textModified = true;
                        }
                    }
                    else
                    {
                        Logger.Warning($"Function not found: {functionName}");
                        textRun.Text = $"[Unknown function: {functionName}]";
                        textModified = true;
                    }
                }
                // 텍스트에 여러 표현식이 있거나 혼합 콘텐츠가 있는 경우
                else
                {
                    // 전체 텍스트를 평가하기 위해 DollarSignEngine 사용
                    var result = _expressionEvaluator.Evaluate(text, variables);
                    textRun.Text = result?.ToString() ?? "";
                    textModified = true;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing PowerPoint function: {text}", ex);
            textRun.Text = $"[Error: {ex.Message}]";
            textModified = true;
        }

        return textModified;
    }

    /// <summary>
    /// PPT 구문 지침에 따라 함수 매개변수 파싱
    /// </summary>
    private string[] ParseFunctionParameters(string parametersString)
    {
        if (string.IsNullOrEmpty(parametersString))
            return Array.Empty<string>();

        var results = new List<string>();
        bool inQuotes = false;
        int currentStart = 0;
        int parenDepth = 0;

        for (int i = 0; i < parametersString.Length; i++)
        {
            char c = parametersString[i];

            // 인용 부호 처리 (따옴표 문자열의 시작/끝)
            if (c == '"' && (i == 0 || parametersString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            // 중첩 괄호 처리
            else if (!inQuotes && c == '(')
            {
                parenDepth++;
            }
            else if (!inQuotes && c == ')')
            {
                parenDepth--;
            }
            // 매개변수 구분자 (최상위 레벨에서만, 인용부호나 중첩 괄호 안에서는 무시)
            else if (c == ',' && !inQuotes && parenDepth == 0)
            {
                results.Add(parametersString.Substring(currentStart, i - currentStart).Trim());
                currentStart = i + 1;
            }
        }

        // 마지막 매개변수 추가
        results.Add(parametersString.Substring(currentStart).Trim());

        // 따옴표 문자열과 명명된 매개변수 정리
        for (int i = 0; i < results.Count; i++)
        {
            var param = results[i].Trim();

            // 명명된 매개변수 처리 (param: value)
            if (param.Contains(":") && !inQuotes)
            {
                var parts = param.Split(new[] { ':' }, 2);
                string paramName = parts[0].Trim();
                string paramValue = parts[1].Trim();

                // 매개변수 값이 따옴표로 묶여 있으면 따옴표 제거
                if (paramValue.StartsWith("\"") && paramValue.EndsWith("\"") && paramValue.Length > 1)
                {
                    paramValue = paramValue.Substring(1, paramValue.Length - 2)
                        .Replace("\\\"", "\"")
                        .Replace("\\\\", "\\")
                        .Replace("\\n", "\n")
                        .Replace("\\r", "\r");
                }

                results[i] = $"{paramName}: {paramValue}";
            }
            // 일반 따옴표 문자열 처리
            else if (param.StartsWith("\"") && param.EndsWith("\"") && param.Length > 1)
            {
                // 따옴표와 이스케이프된 문자 처리
                param = param.Substring(1, param.Length - 2)
                    .Replace("\\\"", "\"")
                    .Replace("\\\\", "\\")
                    .Replace("\\n", "\n")
                    .Replace("\\r", "\r");

                results[i] = param;
            }
        }

        Logger.Debug($"Parsed parameters: {string.Join(", ", results)}");
        return results.ToArray();
    }
}