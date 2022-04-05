using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Specifies possible values for the sections of the table to which the current conditional
/// formatting properties shall be applied when this table style is used.
/// </summary>
public enum TableStyleType
{
    /// <summary>
    /// Specifies that the table formatting applies to odd numbered groupings of rows.
    /// </summary>
    [XmlAttribute("band1Horz")] BandedOddRows,
    /// <summary>
    /// Specifies that the table formatting applies to odd numbered groupings of columns.
    /// </summary>
    [XmlAttribute("band1Vert")] BandedOddColumns,
    /// <summary>
    /// Specifies that the table formatting applies to even numbered groupings of rows.
    /// </summary>
    [XmlAttribute("band2Horz")] BandedEvenRows,
    /// <summary>
    /// Specifies that the table formatting applies to even numbered groupings of columns.
    /// </summary>
    [XmlAttribute("band2Vert")] BandedEvenColumns,
    /// <summary>
    /// Specifies that the table formatting applies to the first column
    /// </summary>
    [XmlAttribute("firstCol")] FirstColumn,
    /// <summary>
    /// Specifies that the table formatting applies to the first row
    /// </summary>
    [XmlAttribute("firstRow")] FirstRow,
    /// <summary>
    /// Specifies that the table formatting applies to the last column
    /// </summary>
    [XmlAttribute("lastCol")] LastColumn,
    /// <summary>
    /// Specifies that the table formatting applies to the last row
    /// </summary>
    [XmlAttribute("lastRow")] LastRow,
    /// <summary>
    /// Specifies that the table formatting applies to the top/right cell
    /// </summary>
    [XmlAttribute("neCell")] TopRightCell,
    /// <summary>
    /// Specifies that the table formatting applies to the top/left cell
    /// </summary>
    [XmlAttribute("nwCell")] TopLeftCell,
    /// <summary>
    /// Specifies that the table formatting applies to the bottom/right cell
    /// </summary>
    [XmlAttribute("seCell")] BottomRightCell,
    /// <summary>
    /// Specifies that the table formatting applies to the bottom/left cell
    /// </summary>
    [XmlAttribute("swCell")] BottomLeftCell,
    /// <summary>
    /// Specifies that the table formatting applies to the entire table
    /// </summary>
    [XmlAttribute("wholeTable")] WholeTable
}