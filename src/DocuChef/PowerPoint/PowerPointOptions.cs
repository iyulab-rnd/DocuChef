namespace DocuChef.PowerPoint;

/// <summary>
/// Options for customizing PowerPoint template processing
/// </summary>
public class PowerPointOptions
{
    /// <summary>
    /// Whether to automatically register built-in functions
    /// </summary>
    public bool RegisterBuiltInFunctions { get; set; } = true;

    /// <summary>
    /// Whether to populate global variables
    /// </summary>
    public bool RegisterGlobalVariables { get; set; } = true;

    /// <summary>
    /// Whether to create new slides when a slide-foreach directive exceeds available slides
    /// </summary>
    public bool CreateNewSlidesWhenNeeded { get; set; } = true;

    /// <summary>
    /// Maximum number of slides that can be generated from a single template slide
    /// </summary>
    public int MaxSlidesFromTemplate { get; set; } = 100;

    /// <summary>
    /// Default image width in pixels when not specified
    /// </summary>
    public int DefaultImageWidth { get; set; } = 300;

    /// <summary>
    /// Default image height in pixels when not specified
    /// </summary>
    public int DefaultImageHeight { get; set; } = 200;

    /// <summary>
    /// Whether to preserve aspect ratio when resizing images by default
    /// </summary>
    public bool PreserveImageAspectRatio { get; set; } = true;

    /// <summary>
    /// Whether to delete temporary files when disposing
    /// </summary>
    public bool CleanupTemporaryFiles { get; set; } = true;

    /// <summary>
    /// Maximum number of iterations for foreach directive
    /// </summary>
    public int MaxIterationItems { get; set; } = 1000;
}