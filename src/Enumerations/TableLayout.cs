using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Specifies the algorithm which shall be used to lay out the contents of this table
/// within the document. When a table is displayed in a document, it can either be displayed
/// using a fixed width or autofit layout algorithm. If ommitted, the value will be assumed "auto".
/// </summary>
public enum TableLayout
{
    /// <summary>
    /// AutoFit Table Layout - This method of table layout uses the preferred widths on the table items
    /// to generate the final sizing of the table, but then uses the contents of each cell to
    /// determine final column widths.
    /// </summary>
    [XmlAttribute("autofit")] AutoFit,

    /// <summary>
    /// Fixed Width Table Layout - This method of table layout uses the preferred widths on the table
    /// items to generate the final sizing of the table, but does not change that size regardless of the
    /// contents of each table cell, hence the table is fixed width.
    /// </summary>
    [XmlAttribute("fixed")] Fixed
}