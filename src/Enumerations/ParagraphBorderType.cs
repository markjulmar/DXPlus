namespace DXPlus;

/// <summary>
/// Border borderEdgeType type
/// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tableborders.aspx
/// </summary>
public enum ParagraphBorderType
{
    /// <summary>
    /// Specifies the border displayed above a set of paragraphs which have the same set of paragraph border settings.
    /// </summary>
    Top,
        
    /// <summary>
    /// Specifies the border displayed to the left of a set of paragraphs which have the same set of paragraph border settings.
    /// </summary>
    Left,
        
    /// <summary>
    /// Specifies the border displayed below a set of paragraphs which have the same set of paragraph border settings.
    /// </summary>
    Bottom,
        
    /// <summary>
    /// Specifies the border displayed to the right of a set of paragraphs which have the same set of paragraph border settings.
    /// </summary>
    Right,
        
    /// <summary>
    /// Specifies the border between each paragraph in a set of paragraphs which have the same set of paragraph border settings.
    /// So if adjoining paragraphs have identical border settings, then there will be one border between them as specified
    /// by the between element. Otherwise the first paragraph will use its bottom border and the following paragraph will use its top border.
    /// </summary>
    Between,
}