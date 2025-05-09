using DocuChef.PowerPoint.DollarSignEngine;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.PowerPoint;

/// <summary>
/// Processes PowerPoint templates with DollarSignEngine for expression evaluation
/// </summary>
internal partial class PowerPointProcessor
{

    private readonly PresentationDocument _document;
    private readonly PowerPointOptions _options;
    private readonly PowerPointContext _context;
    private readonly ExpressionEvaluator _expressionEvaluator;

    /// <summary>
    /// Initializes a new instance of the PowerPointProcessor
    /// </summary>
    public PowerPointProcessor(PresentationDocument document, PowerPointOptions options)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _options = options ?? throw new ArgumentNullException(nameof(options));

        _context = new PowerPointContext
        {
            Options = options
        };

        // Create ExpressionEvaluator for DollarSignEngine integration
        _expressionEvaluator = new ExpressionEvaluator();

        Logger.Debug("PowerPoint processor initialized with DollarSignEngine");
    }

    /// <summary>
    /// Process the PowerPoint template with the provided variables and functions
    /// </summary>
    public void Process(Dictionary<string, object> variables, Dictionary<string, Func<object>> globalVariables, Dictionary<string, PowerPointFunction> functions)
    {
        // Set up context
        _context.Variables = variables ?? new Dictionary<string, object>();
        _context.GlobalVariables = globalVariables ?? new Dictionary<string, Func<object>>();
        _context.Functions = functions ?? new Dictionary<string, PowerPointFunction>();

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

            // Process each slide
            var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
            Logger.Info($"Processing {slideIds.Count} slides");

            for (int i = 0; i < slideIds.Count; i++)
            {
                ProcessSlide(presentationPart, slideIds[i], i);

                // Ensure all changes are saved for this slide
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

            // Save the presentation after processing all slides
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
    /// Process a single slide
    /// </summary>
    private void ProcessSlide(PresentationPart presentationPart, SlideId slideId, int slideIndex)
    {
        Logger.Debug($"Processing slide {slideIndex} with ID {slideId.RelationshipId}");

        // Update slide context
        _context.Slide.Index = slideIndex;
        _context.Slide.Id = slideId.RelationshipId;

        // Get slide part
        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
        if (slidePart == null || slidePart.Slide == null)
        {
            Logger.Warning($"Slide part not found for ID {slideId.RelationshipId}");
            return;
        }

        // Store the SlidePart in context
        _context.SlidePart = slidePart;

        // Get slide notes
        string slideNotes = slidePart.GetNotes();
        _context.Slide.Notes = slideNotes;

        Logger.Debug($"Slide notes: {slideNotes}");

        // Parse directives from slide notes using improved DirectiveParser
        var directives = DirectiveParser.ParseDirectives(slideNotes);

        // Process slide-level directives first (e.g., #slide-foreach)
        foreach (var directive in directives.Where(d => d.Name.StartsWith("slide-")))
        {
            Logger.Debug($"Processing slide directive: {directive.Name}");
            ProcessSlideDirective(presentationPart, slidePart, directive);
        }

        // Process shape directives (e.g., #foreach, #if)
        foreach (var directive in directives.Where(d => !d.Name.StartsWith("slide-")))
        {
            Logger.Debug($"Processing shape directive: {directive.Name}");
            ProcessShapeDirective(slidePart, directive);
        }

        // Process text replacements using DollarSignEngine
        Logger.Debug($"Processing text replacements with DollarSignEngine on slide {slideIndex}");
        ProcessTextReplacements(slidePart);

        // Save slide after processing
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
    /// Prepare variables dictionary combining context variables and global variables
    /// </summary>
    private Dictionary<string, object> PrepareVariablesDictionary()
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
    /// Evaluate an expression using DollarSignEngine
    /// </summary>
    private object EvaluateExpression(string expression)
    {
        // Prepare variables dictionary
        var variables = PrepareVariablesDictionary();

        // Evaluate using DollarSignEngine adapter
        return _expressionEvaluator.Evaluate(expression, variables);
    }

    /// <summary>
    /// Resolve variable value from context
    /// </summary>
    private object ResolveVariableValue(string name)
    {
        // Check for direct variable
        if (_context.Variables.TryGetValue(name, out var value))
            return value;

        // Check for global variable
        if (_context.GlobalVariables.TryGetValue(name, out var factory))
            return factory();

        // Check for property path
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