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
    private readonly TextProcessingHelper _textHelper;

    /// <summary>
    /// Initialize PowerPoint processor
    /// </summary>
    public PowerPointProcessor(PresentationDocument document, PowerPointOptions options)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _options = options ?? throw new ArgumentNullException(nameof(options));

        // Initialize context with options
        _context = new PowerPointContext
        {
            Options = options
        };

        // Create ExpressionEvaluator for DollarSignEngine integration
        _expressionEvaluator = new ExpressionEvaluator();

        // Initialize text processing helper
        _textHelper = new TextProcessingHelper(this, _context);

        Logger.Debug("PowerPoint processor initialized with DollarSignEngine");
    }

    /// <summary>
    /// Process PowerPoint template with variables and functions
    /// </summary>
    public void Process(Dictionary<string, object> variables, Dictionary<string, Func<object>> globalVariables, Dictionary<string, PowerPointFunction> functions)
    {
        // Set up context
        _context.Variables = variables ?? new Dictionary<string, object>();
        _context.GlobalVariables = globalVariables ?? new Dictionary<string, Func<object>>();
        _context.Functions = functions ?? new Dictionary<string, PowerPointFunction>();

        // Add special context variables
        _context.Variables["_context"] = _context;

        try
        {
            // Get all slides in the presentation
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

            // First pass: Analyze and duplicate slides
            var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
            Logger.Info($"Analyzing {slideIds.Count} slides for array references");

            // Create a list to track slides that need duplication
            var slidesToDuplicate = new List<(SlidePart SlidePart, int SlideIndex)>();

            for (int i = 0; i < slideIds.Count; i++)
            {
                var relationshipId = slideIds[i].RelationshipId;
                var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);

                if (slidePart != null)
                {
                    slidesToDuplicate.Add((slidePart, i));
                }
            }

            // Analyze and duplicate slides as needed
            foreach (var (slidePart, slideIndex) in slidesToDuplicate)
            {
                AnalyzeSlideForArrayIndices(presentationPart, slidePart, slideIndex);
            }

            // Second pass: Process all slides including duplicated ones
            slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
            Logger.Info($"Processing {slideIds.Count} slides");

            for (int i = 0; i < slideIds.Count; i++)
            {
                ProcessSlide(presentationPart, slideIds[i], i);
            }

            // Save presentation after processing all slides
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
    /// Prepare variables dictionary combining context variables and global variables
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
}