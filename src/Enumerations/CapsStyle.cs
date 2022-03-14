namespace DXPlus;

/// <summary>
/// Change the caps style of text, for use with Append and AppendLine.
/// </summary>
public enum CapsStyle
{
    /// <summary>
    /// No caps, make all characters are lowercase.
    /// </summary>
    None,

    /// <summary>
    /// All caps, make every character uppercase.
    /// </summary>
    Caps,

    /// <summary>
    /// Small caps, make all characters capital but with a small font size.
    /// </summary>
    SmallCaps
};