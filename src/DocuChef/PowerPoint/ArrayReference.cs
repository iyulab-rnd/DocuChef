namespace DocuChef.PowerPoint;

/// <summary>
/// Represents an array reference found in a PowerPoint slide
/// </summary>
internal class ArrayReference
{
    /// <summary>
    /// The name of the array
    /// </summary>
    public string ArrayName { get; set; }

    /// <summary>
    /// The index referenced in the array
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// The property path after the array index (if any)
    /// </summary>
    public string PropertyPath { get; set; }

    /// <summary>
    /// The full pattern matched in the text
    /// </summary>
    public string Pattern { get; set; }

    /// <summary>
    /// The ID of the shape containing this reference (if available)
    /// </summary>
    public uint? ShapeId { get; set; }

    /// <summary>
    /// The name of the shape containing this reference
    /// </summary>
    public string ShapeName { get; set; }

    /// <summary>
    /// Whether this reference is in an image function
    /// </summary>
    public bool IsInImageFunction { get; set; }
}