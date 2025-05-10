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

        // Check for expression pattern in shape first
        bool hasExpressions = false;
        foreach (var para in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in para.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement != null && !string.IsNullOrEmpty(textElement.Text) &&
                    textElement.Text.Contains("${"))
                {
                    hasExpressions = true;
                    break;
                }
            }
            if (hasExpressions) break;
        }

        if (!hasExpressions)
            return false;

        // Check if expressions span across multiple runs
        bool containsPartialExpressions = ContainsPartialExpressions(shape);

        if (containsPartialExpressions)
        {
            // Handle expressions that span across runs
            return ProcessCrossRunExpressions(shape);
        }
        else
        {
            // Process each run individually (simpler case)
            return ProcessRunsIndividually(shape);
        }
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
    /// Process expressions that span across multiple runs
    /// </summary>
    private bool ProcessCrossRunExpressions(P.Shape shape)
    {
        bool hasChanges = false;

        // Process paragraph by paragraph
        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            // Build complete paragraph text and map runs to positions
            StringBuilder paragraphText = new StringBuilder();
            List<(A.Run Run, int StartPos, int Length)> runMappings = new List<(A.Run, int, int)>();

            foreach (var run in paragraph.Elements<A.Run>().ToList())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement != null && !string.IsNullOrEmpty(textElement.Text))
                {
                    int startPos = paragraphText.Length;
                    string text = textElement.Text;
                    paragraphText.Append(text);
                    runMappings.Add((run, startPos, text.Length));
                }
            }

            // Check if paragraph contains expressions
            string paraText = paragraphText.ToString();
            if (!paraText.Contains("${"))
                continue;

            // Process expressions in complete paragraph text
            string processedText = ProcessExpressions(paraText);
            if (processedText == paraText)
                continue;

            // Map processed text back to runs
            if (MapProcessedTextToRuns(paragraph, runMappings, processedText))
                hasChanges = true;
        }

        return hasChanges;
    }

    /// <summary>
    /// Map processed text back to runs, preserving formatting
    /// </summary>
    private bool MapProcessedTextToRuns(A.Paragraph paragraph, List<(A.Run Run, int StartPos, int Length)> runMappings, string processedText)
    {
        if (runMappings.Count == 0)
            return false;

        // Simple case: If we have just one run, replace its text directly
        if (runMappings.Count == 1)
        {
            var textElement = runMappings[0].Run.GetFirstChild<A.Text>();
            if (textElement != null)
            {
                textElement.Text = processedText;
                return true;
            }
            return false;
        }

        // Complex case: Try to map text back to runs based on relative positions
        // This is an approximate approach that works for simple cases

        // First, clear all existing runs
        foreach (var runInfo in runMappings)
        {
            runInfo.Run.RemoveAllChildren<A.Text>();
        }

        // Determine how to distribute the processed text
        double ratio = (double)processedText.Length / runMappings.Sum(r => r.Length);

        int remainingText = processedText.Length;
        int currentPos = 0;

        for (int i = 0; i < runMappings.Count; i++)
        {
            var runInfo = runMappings[i];

            // Last run gets all remaining text
            if (i == runMappings.Count - 1)
            {
                if (currentPos < processedText.Length)
                {
                    string runText = processedText.Substring(currentPos);
                    runInfo.Run.AppendChild(new A.Text(runText));
                }
                else
                {
                    runInfo.Run.AppendChild(new A.Text(string.Empty));
                }
            }
            else
            {
                // Allocate text proportionally to original length
                int newLength = (int)Math.Ceiling(runInfo.Length * ratio);
                newLength = Math.Min(newLength, remainingText);

                if (newLength > 0 && currentPos < processedText.Length)
                {
                    string runText = processedText.Substring(currentPos, Math.Min(newLength, processedText.Length - currentPos));
                    runInfo.Run.AppendChild(new A.Text(runText));
                    currentPos += runText.Length;
                    remainingText -= runText.Length;
                }
                else
                {
                    runInfo.Run.AppendChild(new A.Text(string.Empty));
                }
            }
        }

        return true;
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

        // Check for array references first (special handling for Items[n].Property patterns)
        var arrayPattern = new Regex(@"\$\{(item|items)(\[\d+\])(\.(\w+))?(:[^}]+)?\}", RegexOptions.IgnoreCase);
        var matches = arrayPattern.Matches(text);

        if (matches.Count > 0)
        {
            foreach (Match match in matches)
            {
                string fullMatch = match.Value;
                try
                {
                    // Normalize array name to 'Items'
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
                    Logger.Warning($"Error evaluating array expression '{fullMatch}': {ex.Message}");
                }
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