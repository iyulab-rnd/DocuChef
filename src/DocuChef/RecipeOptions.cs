using DocuChef.Excel;
using DocuChef.PowerPoint;
using DocuChef.Word;
using System.Globalization;

public class RecipeOptions
{
    /// <summary>
    /// Culture info used for formatting operations
    /// </summary>
    public CultureInfo CultureInfo { get; set; } = CultureInfo.CurrentCulture;

    /// <summary>
    /// String to display when a null value is encountered
    /// </summary>
    public string NullDisplayString { get; set; } = string.Empty;

    /// <summary>
    /// Custom variable resolver for variable handling
    /// </summary>
    public Func<string, Dictionary<string, object>, object?>? VariableResolver { get; set; }

    /// <summary>
    /// Custom formatters for specialized data formatting
    /// </summary>
    public Dictionary<string, Func<object, string, string>> CustomFormatters { get; set; } =
        new Dictionary<string, Func<object, string, string>>();

    /// <summary>
    /// Log callback for library logging
    /// </summary>
    public LogCallback? LogCallback { get; set; }

    /// <summary>
    /// Excel-specific options
    /// </summary>
    public ExcelOptions Excel { get; set; } = new ExcelOptions();

    /// <summary>
    /// Word-specific options
    /// </summary>
    public WordOptions Word { get; set; } = new WordOptions();

    /// <summary>
    /// PowerPoint-specific options
    /// </summary>
    public PowerPointOptions PowerPoint { get; set; } = new PowerPointOptions();
}