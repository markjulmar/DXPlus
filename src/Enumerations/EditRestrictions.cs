namespace DXPlus;

/// <summary>
/// Edit restrictions on the document
/// </summary>
public enum EditRestrictions
{
    /// <summary>
    /// None
    /// </summary>
    None,
        
    /// <summary>
    /// Document is read-only
    /// </summary>
    ReadOnly,
        
    /// <summary>
    /// Forms are editable, but text is read-only
    /// </summary>
    Forms,

    /// <summary>
    /// Comments may be added
    /// </summary>
    Comments,

    /// <summary>
    /// Document is tracking changes
    /// </summary>
    TrackedChanges
}