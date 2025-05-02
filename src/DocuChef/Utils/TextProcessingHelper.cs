using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace DocuChef.Utils;

/// <summary>
/// Helper class for text processing operations across document types
/// </summary>
internal static class TextProcessingHelper
{
    /// <summary>
    /// Process text with variables
    /// </summary>
    public static string ProcessVariables(
        string text,
        Dictionary<string, object> data,
        CultureInfo cultureInfo,
        Func<string, Dictionary<string, object>, object?>? variableResolver = null)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // ClosedXML.Report uses {{variable}} syntax
        // We can return the text as is, as it will be processed by ClosedXML.Report
        return text;
    }

    /// <summary>
    /// Process text with variables using DollarSignEngine
    /// </summary>
    public static async Task<string> ProcessVariablesAsync(
        string text,
        Dictionary<string, object> data,
        CultureInfo cultureInfo,
        bool supportDollarSignSyntax = true,
        Func<string, object, object?>? variableResolver = null,
        IEnumerable<string>? additionalNamespaces = null)
    {
        if (string.IsNullOrEmpty(text)) return text;

        // Skip if no variables
        if (!text.Contains("${") && !text.Contains('{')) return text;

        try
        {
            // Create DollarSignEngine options
            var dollarSignOptions = new DollarSignOption
            {
                SupportDollarSignSyntax = supportDollarSignSyntax,
                FormattingCulture = cultureInfo,
                ThrowOnMissingParameter = false,
                VariableResolver = (expr, param) => {
                    if (variableResolver != null)
                    {
                        return variableResolver(expr, data);
                    }
                    return null;
                }
            };

            // Add additional namespaces
            if (additionalNamespaces != null && additionalNamespaces.Any())
            {
                dollarSignOptions.AdditionalNamespaces = [.. additionalNamespaces];
            }

            // Process the text using DollarSignEngine
            return await DollarSign.EvalAsync(text, data, dollarSignOptions);
        }
        catch (Exception ex)
        {
            throw new TemplateException($"Error processing text with variables: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Process text with handlebars templates (Excel style)
    /// </summary>
    public static string ProcessHandlebarsTemplate(
        string text,
        Dictionary<string, object> contextData,
        Func<string, Dictionary<string, object>, object?>? variableResolver,
        CultureInfo cultureInfo,
        string nullDisplayString,
        Dictionary<string, Func<object, string, string>>? customFormatters)
    {
        if (string.IsNullOrEmpty(text)) return text;

        // Skip if no variables
        if (!text.Contains("{{")) return text;

        // Regex patterns for handlebars syntax
        var ifRegex = new Regex(@"\{\{#if\s+([^{}]+)\}\}(.*?)(\{\{else\}\}(.*?))?\{\{\/if\}\}", RegexOptions.Compiled | RegexOptions.Singleline);
        var eachRegex = new Regex(@"\{\{#each\s+([^{}]+)\}\}(.*?)\{\{\/each\}\}", RegexOptions.Compiled | RegexOptions.Singleline);
        var variableRegex = new Regex(@"\{\{([^{}]+)\}\}", RegexOptions.Compiled);

        // 인식된 루프 변수 추적을 위한 컨텍스트 정보
        var loopVariables = new Dictionary<string, object>();

        // Process each blocks first
        text = eachRegex.Replace(text, match => {
            var collectionPath = match.Groups[1].Value.Trim();
            var template = match.Groups[2].Value;

            var collection = ResolveCollection(collectionPath, contextData, variableResolver);
            if (collection == null || !collection.Any())
                return string.Empty;

            var result = new System.Text.StringBuilder();

            foreach (var item in collection)
            {
                // 루프 컨텍스트 생성 (현재 항목 + 글로벌 컨텍스트)
                var loopContext = CreateCombinedContext(item, contextData);
                loopContext["item"] = item; // 현재 항목을 item으로 명시적 추가

                // 하위 템플릿 처리
                var processed = ProcessHandlebarsTemplate(
                    template,
                    loopContext,
                    variableResolver,
                    cultureInfo,
                    nullDisplayString,
                    customFormatters);

                result.Append(processed);
            }

            return result.ToString();
        });

        // Process conditional blocks
        text = ifRegex.Replace(text, match => {
            var condition = match.Groups[1].Value.Trim();
            var trueContent = match.Groups[2].Value;
            var falseContent = match.Groups.Count > 4 ? match.Groups[4].Value : string.Empty;

            bool result = EvaluateCondition(condition, contextData, variableResolver);

            // 조건에 맞는 텍스트 처리
            var selectedContent = result ? trueContent : falseContent;

            // 중첩된 핸들바 처리
            return ProcessHandlebarsTemplate(
                selectedContent,
                contextData,
                variableResolver,
                cultureInfo,
                nullDisplayString,
                customFormatters);
        });

        // Process variables
        return variableRegex.Replace(text, match => {
            var expression = match.Groups[1].Value.Trim();

            // Extract format specifier if present
            string? format = null;
            if (expression.Contains(':'))
            {
                var parts = expression.Split(':');
                expression = parts[0].Trim();
                format = parts[1].Trim();
            }

            // First check loop variables if any
            if (loopVariables.TryGetValue(expression, out var loopValue))
            {
                return FormatValue(loopValue, format, cultureInfo, nullDisplayString, customFormatters);
            }

            // Try to resolve from context
            var value = ResolveVariable(expression, contextData, variableResolver);
            return FormatValue(value, format, cultureInfo, nullDisplayString, customFormatters);
        });
    }

    /// <summary>
    /// Format a value with the specified format
    /// </summary>
    private static string FormatValue(
        object? value,
        string? format,
        CultureInfo cultureInfo,
        string nullDisplayString,
        Dictionary<string, Func<object, string, string>>? customFormatters)
    {
        // 널 값 처리
        if (value == null)
            return nullDisplayString;

        // 커스텀 포맷터 적용
        if (!string.IsNullOrEmpty(format) && customFormatters != null && customFormatters.TryGetValue(format, out var formatter))
        {
            return formatter(value, format);
        }

        // 기본 형식 처리
        if (string.IsNullOrEmpty(format))
            return $"{value}";

        // 표준 형식 지정자 처리
        if (value is IFormattable formattable)
            return formattable.ToString(format, cultureInfo);

        // 기본 문자열 변환
        return $"{value}";
    }

    /// <summary>
    /// Evaluate a condition based on variable resolution
    /// </summary>
    public static bool EvaluateCondition(
        string condition,
        Dictionary<string, object> data,
        Func<string, Dictionary<string, object>, object?>? variableResolver = null)
    {
        // Try to get direct boolean value
        if (condition.Equals("true", StringComparison.OrdinalIgnoreCase))
            return true;

        if (condition.Equals("false", StringComparison.OrdinalIgnoreCase))
            return false;

        // item. 접두사 처리 - 루프 컨텍스트의 경우
        if (condition.StartsWith("item.", StringComparison.OrdinalIgnoreCase) && data.ContainsKey("item"))
        {
            var itemObj = data["item"];
            var propName = condition.Substring(5); // "item."을 제외한 속성 이름

            if (itemObj != null)
            {
                // item 객체에서 속성 값 추출
                var property = itemObj.GetType().GetProperty(propName,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

                if (property != null)
                {
                    var propValue = property.GetValue(itemObj);
                    return IsTruthy(propValue);
                }
            }
        }

        // 비교 연산 처리 (간단한 비교만 지원)
        if (condition.Contains('='))
        {
            var parts = condition.Split('=', 2);
            if (parts.Length == 2)
            {
                var leftValue = ResolveVariable(parts[0].Trim(), data, variableResolver);
                var rightValue = parts[1].Trim().Trim('"', '\''); // 리터럴 문자열 따옴표 제거

                // 문자열 비교
                return string.Equals(leftValue?.ToString(), rightValue, StringComparison.OrdinalIgnoreCase);
            }
        }

        // Try to resolve the variable
        var value = ResolveVariable(condition, data, variableResolver);
        return IsTruthy(value);
    }

    /// <summary>
    /// Determine if a value is "truthy"
    /// </summary>
    private static bool IsTruthy(object? value)
    {
        return value switch
        {
            bool b => b,
            string s => !string.IsNullOrEmpty(s),
            int i => i != 0,
            double d => d != 0,
            decimal d => d != 0,
            null => false,
            _ => true  // Non-null values are truthy
        };
    }

    /// <summary>
    /// Resolve a variable from context data
    /// </summary>
    public static object? ResolveVariable(
        string expression,
        Dictionary<string, object> contextData,
        Func<string, Dictionary<string, object>, object?>? variableResolver = null)
    {
        // Try custom variable resolver first
        if (variableResolver != null)
        {
            var result = variableResolver(expression, contextData);
            if (result != null)
                return result;
        }

        // Try direct access to the dictionary
        if (contextData.TryGetValue(expression, out var value))
            return value;

        // Handle dotted path (nested properties)
        if (expression.Contains('.'))
        {
            var parts = expression.Split('.');
            object? current = contextData;

            foreach (var part in parts)
            {
                if (current == null)
                    return null;

                // Handle dictionary
                if (current is Dictionary<string, object> dict)
                {
                    if (!dict.TryGetValue(part, out current))
                        return null;
                    continue;
                }

                // Handle general dictionary
                if (current is IDictionary genDict)
                {
                    if (genDict.Contains(part))
                    {
                        current = genDict[part];
                        continue;
                    }
                    return null;
                }

                // Handle objects with properties
                var property = current.GetType().GetProperty(part,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (property == null)
                    return null;

                current = property.GetValue(current);
            }

            return current;
        }

        return null;
    }

    /// <summary>
    /// Resolve a collection path from data
    /// </summary>
    public static IEnumerable<object> ResolveCollection(
        string path,
        Dictionary<string, object> data,
        Func<string, Dictionary<string, object>, object?>? variableResolver = null)
    {
        var value = ResolveVariable(path, data, variableResolver);
        if (value == null)
            return Enumerable.Empty<object>();

        // If already an enumerable, return it
        if (value is IEnumerable<object> typedCollection)
            return typedCollection;

        // Handle non-generic enumerable
        if (value is IEnumerable genericCollection && value is not string)
        {
            return genericCollection.Cast<object>();
        }

        return Enumerable.Empty<object>();
    }

    /// <summary>
    /// Create a combined data context for collection item processing
    /// </summary>
    public static Dictionary<string, object> CreateCombinedContext(
        object item,
        Dictionary<string, object> globalData)
    {
        // Convert item properties to dictionary
        var itemData = DataConverter.ObjectToDictionary(item);
        var combinedData = new Dictionary<string, object>(globalData, StringComparer.OrdinalIgnoreCase)
        {
            // Add item as "item" variable
            ["item"] = item,
            ["this"] = item
        };

        // Add item properties to root level (optional)
        foreach (var kvp in itemData)
        {
            if (!combinedData.ContainsKey(kvp.Key))
            {
                combinedData[kvp.Key] = kvp.Value;
            }
        }

        return combinedData;
    }
}