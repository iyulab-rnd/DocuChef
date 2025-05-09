using System.Text.RegularExpressions;
using DollarSignEngine;
using System.Globalization;

namespace DocuChef.PowerPoint.DollarSignEngine;

/// <summary>
/// PowerPoint 템플릿의 표현식을 처리하기 위한 DollarSignEngine 어댑터
/// </summary>
internal class ExpressionEvaluator
{
    private readonly DollarSignOption _options;
    private readonly CultureInfo _cultureInfo;

    /// <summary>
    /// ExpressionEvaluator 인스턴스 초기화
    /// </summary>
    public ExpressionEvaluator(CultureInfo cultureInfo = null)
    {
        _cultureInfo = cultureInfo ?? CultureInfo.CurrentCulture;

        _options = new DollarSignOption
        {
            SupportDollarSignSyntax = true,  // PPT 구문 지침에 따라 ${variable} 구문 사용
            ThrowOnMissingParameter = false, // 누락된 매개변수에 대해 예외를 발생시키지 않고 대신 자리 표시자 표시
            EnableDebugLogging = false,      // 기본적으로 디버그 로깅 비활성화
            PreferCallbackResolution = true, // 특수 함수에 콜백 해결 선호
            VariableResolver = HandleSpecialResolution // PowerPoint 함수 및 배열 인덱스 처리를 위한 사용자 지정 해결기
        };
    }

    /// <summary>
    /// 표현식을 동기적으로 평가
    /// </summary>
    public object Evaluate(string expression, Dictionary<string, object> variables)
    {
        try
        {
            // 1. 포맷 지정자가 있는 표현식 특별 처리
            // 예: ${item[0].Price:C0}
            var formatMatch = Regex.Match(expression, @"^\${([\w\[\]\.]+):([^}]+)}$");
            if (formatMatch.Success)
            {
                return EvaluateFormattedExpression(formatMatch.Groups[1].Value, formatMatch.Groups[2].Value, variables);
            }

            // 2. 배열 인덱스가 있는 표현식 특별 처리
            // 예: ${item[0].Id} 또는 item[0].Id
            var arrayMatch = Regex.Match(expression, @"^\${?([\w]+)\[(\d+)\](\.(\w+))?}?$");
            if (arrayMatch.Success)
            {
                return EvaluateArrayIndexExpression(arrayMatch, variables);
            }

            // 3. PowerPoint 특수 함수 처리
            if (expression.StartsWith("ppt."))
            {
                return EvaluateAsync(expression, variables).GetAwaiter().GetResult();
            }

            // 4. ${...} 구문이 이미 포함된 경우 직접 평가
            if (expression.Contains("${"))
            {
                Logger.Debug($"Evaluating text with embedded variables: {expression}");
                var result = DollarSign.EvalAsync(expression, variables, _options).GetAwaiter().GetResult();
                Logger.Debug($"Expression result: {result}");
                return result;
            }

            // 5. ${...}로 시작하지 않는 경우 평가를 위해 래핑
            if (!expression.StartsWith("${"))
            {
                expression = "${" + expression + "}";
            }

            // DollarSignEngine을 사용하여 평가
            var evalResult = DollarSign.EvalAsync(expression, variables, _options).GetAwaiter().GetResult();
            Logger.Debug($"Expression result: {evalResult}");
            return evalResult;
        }
        catch (Exception ex)
        {
            if (ex is Microsoft.CodeAnalysis.Scripting.CompilationErrorException compEx)
            {
                // 컴파일 오류 자세히 로깅
                Logger.Error($"Compilation error in expression: {expression}");
                foreach (var diagnostic in compEx.Diagnostics)
                {
                    Logger.Error($"  - {diagnostic.GetMessage()}");
                }
            }

            Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// 포맷 지정자가 있는 표현식 평가
    /// </summary>
    private object EvaluateFormattedExpression(string valueExpr, string format, Dictionary<string, object> variables)
    {
        try
        {
            Logger.Debug($"Evaluating formatted expression: value={valueExpr}, format={format}");

            // 변수 또는 표현식 값 가져오기
            object value;

            // 단순 변수 참조인 경우
            if (variables.TryGetValue(valueExpr, out value))
            {
                Logger.Debug($"Found direct variable: {valueExpr} = {value}");
            }
            else
            {
                // 복합 표현식인 경우 평가
                string wrappedExpr = "${" + valueExpr + "}";
                value = DollarSign.EvalAsync(wrappedExpr, variables, _options).GetAwaiter().GetResult();
                Logger.Debug($"Evaluated expression: {valueExpr} = {value}");
            }

            if (value == null)
                return "";

            // 포맷 적용
            if (value is IFormattable formattable)
            {
                try
                {
                    // 문화권 설정으로 포맷 적용
                    string formattedValue = formattable.ToString(format, _cultureInfo);
                    Logger.Debug($"Formatted value: {formattedValue}");
                    return formattedValue;
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Error applying format '{format}' to value '{value}': {ex.Message}");
                }
            }

            // 기본 문자열 반환
            return value.ToString();
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating formatted expression '{valueExpr}:{format}': {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// 배열 인덱스 표현식 평가
    /// </summary>
    private object EvaluateArrayIndexExpression(Match match, Dictionary<string, object> variables)
    {
        string arrayName = match.Groups[1].Value; // item
        int index = int.Parse(match.Groups[2].Value); // 0
        string propPath = match.Groups[4].Success ? match.Groups[4].Value : null; // Id or null

        Logger.Debug($"Evaluating array index expression: array={arrayName}, index={index}, prop={propPath}");

        // 변수 딕셔너리에서 배열 찾기
        if (!variables.TryGetValue(arrayName, out var arrayObj) || arrayObj == null)
        {
            Logger.Warning($"Array or collection '{arrayName}' not found");
            return "";
        }

        // 배열에서 항목 가져오기
        object item = null;

        try
        {
            if (arrayObj is IList list)
            {
                if (index >= 0 && index < list.Count)
                {
                    item = list[index];
                }
                else
                {
                    Logger.Warning($"Index {index} is out of range for list with {list.Count} items");
                    return "";
                }
            }
            else if (arrayObj is Array array)
            {
                if (index >= 0 && index < array.Length)
                {
                    item = array.GetValue(index);
                }
                else
                {
                    Logger.Warning($"Index {index} is out of range for array with length {array.Length}");
                    return "";
                }
            }
            else
            {
                // 다른 인덱싱 가능 컬렉션 처리 시도
                var indexerProp = arrayObj.GetType().GetProperty("Item");
                if (indexerProp != null)
                {
                    try
                    {
                        item = indexerProp.GetValue(arrayObj, new object[] { index });
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to access indexer: {ex.Message}");
                        return "";
                    }
                }
                else
                {
                    Logger.Warning($"Object of type {arrayObj.GetType().Name} does not support indexing");
                    return "";
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error accessing item at index {index}: {ex.Message}", ex);
            return "";
        }

        if (item == null)
        {
            Logger.Warning($"Item at index {index} is null");
            return "";
        }

        // 속성 경로가 없는 경우 항목 자체 반환
        if (string.IsNullOrEmpty(propPath))
        {
            return item.ToString();
        }

        // 속성 값 가져오기
        try
        {
            var property = item.GetType().GetProperty(propPath);
            if (property != null)
            {
                var propValue = property.GetValue(item);
                return propValue?.ToString() ?? "";
            }
            else
            {
                Logger.Warning($"Property '{propPath}' not found on type {item.GetType().Name}");
                return "";
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error accessing property '{propPath}': {ex.Message}", ex);
            return "";
        }
    }

    /// <summary>
    /// 표현식을 비동기적으로 평가
    /// </summary>
    private async Task<object> EvaluateAsync(string expression, Dictionary<string, object> variables)
    {
        try
        {
            Logger.Debug($"Evaluating expression: {expression}");

            // PowerPoint 특수 함수인 경우 (ppt.)
            if (expression.StartsWith("ppt."))
            {
                return await HandlePptFunctionAsync(expression, variables);
            }

            // 이미 ${...}가 포함된 경우 직접 평가
            if (expression.Contains("${"))
            {
                Logger.Debug($"Evaluating text with embedded variables: {expression}");
                var result = await DollarSign.EvalAsync(expression, variables, _options);
                Logger.Debug($"Expression result: {result}");
                return result;
            }

            // ${...}로 시작하지 않는 경우 평가를 위해 래핑
            if (!expression.StartsWith("${"))
            {
                expression = "${" + expression + "}";
            }

            // DollarSignEngine을 사용하여 평가
            var evalResult = await DollarSign.EvalAsync(expression, variables, _options);
            Logger.Debug($"Expression result: {evalResult}");
            return evalResult;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{expression}'", ex);
            throw new DocuChefException($"Error evaluating expression '{expression}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// PowerPoint 특수 함수 및 배열 인덱스를 위한 사용자 지정 해결기
    /// </summary>
    private object HandleSpecialResolution(string expression, object parameters)
    {
        Logger.Debug($"Special resolution for expression: {expression}");

        // "ppt."로 시작하는 경우 PowerPoint 함수로 처리
        if (expression.StartsWith("ppt."))
        {
            // 매개변수를 사전으로 변환
            Dictionary<string, object> variables;
            if (parameters is Dictionary<string, object> dict)
            {
                variables = dict;
            }
            else
            {
                variables = new Dictionary<string, object>();
                if (parameters != null)
                {
                    var props = parameters.GetType().GetProperties();
                    foreach (var prop in props)
                    {
                        if (prop.CanRead)
                        {
                            try
                            {
                                var value = prop.GetValue(parameters);
                                variables[prop.Name] = value;
                            }
                            catch
                            {
                                // 예외를 발생시키는 속성 건너뛰기
                            }
                        }
                    }
                }
            }

            // PowerPoint 함수 처리
            var task = HandlePptFunctionAsync(expression, variables);
            return task.GetAwaiter().GetResult();
        }

        // 배열 인덱스 표현식 처리: item[0].Property
        var arrayMatch = Regex.Match(expression, @"^([\w]+)\[(\d+)\](\.(\w+))?$");
        if (arrayMatch.Success)
        {
            Dictionary<string, object> variables;
            if (parameters is Dictionary<string, object> dict)
            {
                variables = dict;
            }
            else
            {
                variables = new Dictionary<string, object>();
                // 필요한 경우 parameters에서 변수 추출
            }

            return EvaluateArrayIndexExpression(arrayMatch, variables);
        }

        // DollarSignEngine이 표준 표현식을 처리하도록 null 반환
        return null;
    }

    /// <summary>
    /// PowerPoint 특수 함수 처리 (ppt.Image, ppt.Chart, ppt.Table)
    /// </summary>
    private async Task<object> HandlePptFunctionAsync(string expression, Dictionary<string, object> variables)
    {
        Logger.Debug($"Handling PowerPoint function: {expression}");

        // 함수 표현식 파싱: ppt.Function("arg", param1: value1, param2: value2)
        var match = Regex.Match(expression, @"ppt\.(\w+)\((.+)\)");
        if (!match.Success)
        {
            Logger.Warning($"Invalid PowerPoint function format: {expression}");
            return $"[Invalid function: {expression}]";
        }

        string functionName = match.Groups[1].Value;
        string argsString = match.Groups[2].Value;

        Logger.Debug($"Function: {functionName}, Args: {argsString}");

        // 인용 문자열과 명명된 매개변수의 적절한 처리로 인수 파싱
        var args = ParseFunctionArguments(argsString);

        // 인수에 표현식이 포함된 경우 DollarSign 엔진을 사용하여 해결
        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            // 인수에 ${...}가 포함되어 있거나 param:${value}와 같은 명명된 매개변수인 경우
            if (arg.Contains("${") || (arg.Contains(":") && arg.Split(':', 2)[1].Contains("${")))
            {
                // 명명된 매개변수의 경우 (param:value), 값 부분만 평가
                if (arg.Contains(":"))
                {
                    var parts = arg.Split(':', 2);
                    string paramName = parts[0].Trim();
                    string paramValue = parts[1].Trim();

                    // 매개변수 값 평가
                    var resolvedValue = await DollarSign.EvalAsync(paramValue, variables, _options);
                    args[i] = $"{paramName}: {resolvedValue}";
                }
                else
                {
                    // 인수의 직접 평가
                    var resolvedValue = await DollarSign.EvalAsync(arg, variables, _options);
                    args[i] = resolvedValue?.ToString() ?? string.Empty;
                }
            }
        }

        // PowerPoint 함수 검색 및 실행
        if (variables.TryGetValue($"ppt.{functionName}", out var funcObj) &&
            funcObj is PowerPointFunction function)
        {
            Logger.Debug($"Found registered PowerPoint function: {functionName}");

            // PowerPoint 컨텍스트 검색
            PowerPointContext context = null;
            if (variables.TryGetValue("_context", out var ctxObj) &&
                ctxObj is PowerPointContext ctx)
            {
                context = ctx;
            }
            else
            {
                // 컨텍스트가 없는 경우 (이 경우는 드물지만) 임시 컨텍스트 생성
                context = new PowerPointContext
                {
                    Variables = new Dictionary<string, object>(variables)
                };
            }

            try
            {
                // 함수 실행
                var result = function.Execute(context, null, args);
                Logger.Debug($"Function result: {result}");
                return result;
            }
            catch (Exception ex)
            {
                Logger.Error($"Error executing function {functionName}: {ex.Message}", ex);
                return $"[Error in {functionName}: {ex.Message}]";
            }
        }

        Logger.Warning($"Function not found: ppt.{functionName}");
        return $"[Unknown function: ppt.{functionName}]";
    }

    /// <summary>
    /// 인용 문자열과 명명된 매개변수의 개선된 처리로 함수 인수 파싱
    /// </summary>
    private string[] ParseFunctionArguments(string argsString)
    {
        if (string.IsNullOrEmpty(argsString))
            return Array.Empty<string>();

        var args = new List<string>();
        bool inQuotes = false;
        int start = 0;
        int parenthesesDepth = 0;

        for (int i = 0; i < argsString.Length; i++)
        {
            char c = argsString[i];

            if (c == '"')
            {
                // 이스케이프된 따옴표 처리
                if (i > 0 && argsString[i - 1] == '\\')
                {
                    continue;
                }
                inQuotes = !inQuotes;
            }
            else if (c == '(' && !inQuotes)
            {
                parenthesesDepth++;
            }
            else if (c == ')' && !inQuotes)
            {
                parenthesesDepth--;
            }
            else if (c == ',' && !inQuotes && parenthesesDepth == 0)
            {
                // 인수 끝
                args.Add(argsString.Substring(start, i - start).Trim());
                start = i + 1;
            }
        }

        // 마지막 인수 추가
        if (start < argsString.Length)
        {
            args.Add(argsString.Substring(start).Trim());
        }

        // 따옴표 정리 및 명명된 매개변수 처리
        for (int i = 0; i < args.Count; i++)
        {
            string arg = args[i].Trim();

            // 명명된 매개변수 처리 (param: value)
            if (arg.Contains(":") && !inQuotes)
            {
                var parts = arg.Split(new[] { ':' }, 2);
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

                arg = $"{paramName}: {paramValue}";
            }
            // 일반 따옴표 문자열 처리
            else if (arg.StartsWith("\"") && arg.EndsWith("\"") && arg.Length > 1)
            {
                arg = arg.Substring(1, arg.Length - 2)
                    .Replace("\\\"", "\"")
                    .Replace("\\\\", "\\")
                    .Replace("\\n", "\n")
                    .Replace("\\r", "\r");
            }

            args[i] = arg;
        }

        return args.ToArray();
    }
}