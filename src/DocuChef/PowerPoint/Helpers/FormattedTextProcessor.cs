using System.Text;
using System.Text.RegularExpressions;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Helper class for processing formatted text with expression evaluation and formatting preservation
/// </summary>
internal class FormattedTextProcessor
{
    private readonly PowerPointProcessor _processor;
    private readonly Dictionary<string, object> _variables;
    private static readonly Regex _expressionPattern = new Regex(@"\$\{([^{}]+)\}", RegexOptions.Compiled);
    private static readonly Regex _partialExpressionPattern = new Regex(@"\$\{|\}", RegexOptions.Compiled);

    /// <summary>
    /// Initialize formatted text processor
    /// </summary>
    public FormattedTextProcessor(PowerPointProcessor processor, Dictionary<string, object> variables)
    {
        _processor = processor;
        _variables = variables ?? new Dictionary<string, object>();
    }

    /// <summary>
    /// Process a shape with formatting preservation
    /// </summary>
    public bool ProcessShapeTextWithFormatting(P.Shape shape)
    {
        if (shape?.TextBody == null)
            return false;

        bool hasChanges = false;

        // Check if any run in the shape contains partial expressions
        bool containsPartialExpressions = ContainsPartialExpressions(shape);
        if (!containsPartialExpressions)
        {
            // No partial expressions, process each run individually
            return ProcessRunsIndividually(shape);
        }

        // If we have partial expressions, handle special cases for "item[n].Property" patterns
        return ProcessArrayExpressions(shape);
    }

    /// <summary>
    /// Process each run individually for expression evaluation
    /// </summary>
    private bool ProcessRunsIndividually(P.Shape shape)
    {
        bool hasChanges = false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            foreach (var run in paragraph.Elements<A.Run>().ToList())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string originalText = textElement.Text;
                if (!originalText.Contains("${"))
                    continue;

                // Process expressions in this run
                string processedText = ProcessExpressions(originalText);
                if (processedText == originalText)
                    continue;

                // Update the text
                textElement.Text = processedText;
                hasChanges = true;
            }
        }

        return hasChanges;
    }

    /// <summary>
    /// Process array expressions that may be split across multiple runs
    /// </summary>
    private bool ProcessArrayExpressions(P.Shape shape)
    {
        bool hasChanges = false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            // Pattern to detect item[n].Property expressions
            var itemPattern = new Regex(@"\$\{(item|items)(\[\d+\])(\.(\w+))?", RegexOptions.IgnoreCase);
            var runs = paragraph.Elements<A.Run>().ToList();

            // Skip if no runs
            if (runs.Count == 0)
                continue;

            // For each run, check if it contains the start of an array expression
            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;

                // Check if this run has a partial array expression
                var match = itemPattern.Match(text);
                if (!match.Success)
                {
                    // If it has a complete expression, process it normally
                    if (text.Contains("${") && text.Contains("}"))
                    {
                        string processed = ProcessExpressions(text);
                        if (processed != text)
                        {
                            textElement.Text = processed;
                            hasChanges = true;
                        }
                    }
                    continue;
                }

                // Found start of array expression, collect all text parts
                string arrayName = match.Groups[1].Value;
                bool foundComplete = false;
                StringBuilder completePart = new StringBuilder(text);

                // Check next runs to build complete expression
                for (int j = i + 1; j < runs.Count && !foundComplete; j++)
                {
                    var nextRun = runs[j];
                    var nextText = nextRun.GetFirstChild<A.Text>()?.Text;
                    if (string.IsNullOrEmpty(nextText))
                        continue;

                    completePart.Append(nextText);
                    string combined = completePart.ToString();

                    // Check if we have a complete expression now
                    int openBrace = combined.IndexOf("${");
                    int closeBrace = combined.IndexOf("}", openBrace);

                    if (closeBrace > openBrace)
                    {
                        foundComplete = true;
                        string completeExpr = combined.Substring(openBrace, closeBrace - openBrace + 1);

                        // Process the complete expression
                        string processed = ProcessExpressions(completeExpr);
                        if (processed != completeExpr)
                        {
                            // Replace in first run
                            textElement.Text = text.Replace(match.Value, processed);
                            hasChanges = true;

                            // Remove expression parts from other runs
                            for (int k = i + 1; k <= j; k++)
                            {
                                var runToModify = runs[k];
                                var textToModify = runToModify.GetFirstChild<A.Text>();
                                if (textToModify != null)
                                {
                                    string runText = textToModify.Text;
                                    int endPos = k == j ? runText.IndexOf("}") + 1 : runText.Length;
                                    if (endPos > 0)
                                    {
                                        textToModify.Text = runText.Substring(endPos);
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }

        return hasChanges;
    }

    /// <summary>
    /// Check if the shape contains partial expressions (expressions split across runs)
    /// </summary>
    private bool ContainsPartialExpressions(P.Shape shape)
    {
        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            bool hasOpenBrace = false;
            bool hasCloseBrace = false;

            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;

                // Check for complete expressions in this run
                if (text.Contains("${") && text.Contains("}"))
                {
                    // Has complete expression, continue to next run
                    continue;
                }

                // Check for partial expressions
                if (text.Contains("${"))
                {
                    hasOpenBrace = true;
                }
                if (text.Contains("}"))
                {
                    if (hasOpenBrace)
                    {
                        // Found a split expression
                        return true;
                    }
                    hasCloseBrace = true;
                }
            }

            // If we have an open brace without a close brace, it's split
            if (hasOpenBrace && !hasCloseBrace)
                return true;
        }

        return false;
    }

    /// <summary>
    /// Process expressions in text
    /// </summary>
    private string ProcessExpressions(string text)
    {
        if (string.IsNullOrEmpty(text) || !text.Contains("${"))
            return text;

        // Check for simple item[n].property replacements
        var itemPattern = new Regex(@"\$\{(item|items)(\[\d+\])(\.(\w+))?(:[^}]+)?\}", RegexOptions.IgnoreCase);
        var matches = itemPattern.Matches(text);

        foreach (Match match in matches)
        {
            string fullMatch = match.Value;
            try
            {
                // Normalize to Items for consistency
                string normalizedExpr = fullMatch.Replace(match.Groups[1].Value, "Items");

                // Evaluate the expression
                var result = _processor.EvaluateCompleteExpression(normalizedExpr, _variables);
                if (result != null)
                {
                    text = text.Replace(fullMatch, result.ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error evaluating expression '{fullMatch}': {ex.Message}");
            }
        }

        // Process other expressions
        return _expressionPattern.Replace(text, match =>
        {
            try
            {
                var result = _processor.EvaluateCompleteExpression(match.Value, _variables);
                return result?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error evaluating expression '{match.Value}': {ex.Message}");
                return match.Value;
            }
        });
    }
}