namespace DXPlus;

/// <summary>
/// Table Border Types
/// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tableborders.aspx
/// </summary>
public enum TableBorderType
{
    /// <summary>
    /// Specifies the border displayed above a table.
    /// </summary>
    Top,

    /// <summary>
    /// Specifies the border displayed to the left of a table.
    /// </summary>
    Left,

    /// <summary>
    /// Specifies the border displayed below a table.
    /// </summary>
    Bottom,

    /// <summary>
    /// Specifies the border displayed to the right of a table.
    /// </summary>
    Right,

    /// <summary>
    /// Specifies the border displayed on the inside horizontal edge.
    /// </summary>
    InsideH,

    /// <summary>
    /// Specifies the border displayed on the inside vertical edge.
    /// </summary>
    InsideV
}