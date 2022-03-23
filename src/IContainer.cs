namespace DXPlus;

/// <summary>
/// This interface provides basic methods to insert/add/remove items from a FirstParagraph, TableCell, Header or Footer.
/// </summary>
public interface IContainer
{
    /// <summary>
    /// Enumerate all blocks in the container.
    /// </summary>
    IEnumerable<Block> Blocks { get; }

    /// <summary>
    /// Enumerate all paragraphs inside this container.
    /// </summary>
    IEnumerable<Paragraph> Paragraphs { get; }

    /// <summary>
    /// Enumerate all sections associated with this container.
    /// </summary>
    IEnumerable<Section> Sections { get; }

    /// <summary>
    /// Retrieve all Table objects in the container
    /// </summary>
    IEnumerable<Table> Tables { get; }

    /// <summary>
    /// Retrieve a list of all hyperlinks in the document
    /// </summary>
    IEnumerable<Hyperlink> Hyperlinks { get; }

    /// <summary>
    /// Retrieve a list of all drawings in the document
    /// </summary>
    IEnumerable<Drawing> Drawings { get; }

    /// <summary>
    /// Removes paragraph at specified position
    /// </summary>
    /// <param name="index">Index of paragraph to remove</param>
    /// <returns>True if removed</returns>
    bool RemoveAt(int index);

    /// <summary>
    /// Removes paragraph
    /// </summary>
    /// <param name="paragraph">FirstParagraph to remove</param>
    /// <returns>True if removed</returns>
    bool Remove(Paragraph paragraph);

    /// <summary>
    /// Replace matched text with a new value
    /// </summary>
    /// <param name="findText">Text value to search for</param>
    /// <param name="replaceText">Replacement value</param>
    /// <param name="comparisonType">Comparison type - defaults to CurrentCulture</param>
    bool FindReplace(string findText, string? replaceText, StringComparison comparisonType = StringComparison.CurrentCulture);

    /// <summary>
    /// Insert a text block at a specific bookmark
    /// </summary>
    /// <param name="bookmarkName">Bookmark name</param>
    /// <param name="toInsert">Text to insert</param>
    bool InsertTextAtBookmark(string bookmarkName, string toInsert);

    /// <summary>
    /// Insert a paragraph into this container at a specific index
    /// </summary>
    /// <param name="index">Index to insert into</param>
    /// <param name="paragraph">FirstParagraph to insert</param>
    /// <returns>Inserted paragraph</returns>
    Paragraph Insert(int index, Paragraph paragraph);

    /// <summary>
    /// Add a paragraph at the end of the container
    /// </summary>
    Paragraph Add(Paragraph paragraph);

    /// <summary>
    /// Add a new section to the container
    /// </summary>
    Section AddSection(SectionBreakType breakType = SectionBreakType.NextPage);

    /// <summary>
    /// Add a new page break to the container
    /// </summary>
    void AddPageBreak();

    /// <summary>
    /// Create a set of paragraphs from a series of strings. Each string will add a new paragraph
    /// to the document.
    /// </summary>
    /// <param name="paragraphs">Text to add</param>
    /// <returns>Last paragraph added, or null if no paragraphs were added</returns>
    Paragraph? AddRange(IEnumerable<string> paragraphs);

    /// <summary>
    /// Create a set of paragraphs from a series of strings. Each string will add a new paragraph
    /// to the document.
    /// </summary>
    /// <param name="paragraphs">Paragraphs to add</param>
    /// <returns>Last paragraph added, or null if no paragraphs were added</returns>
    Paragraph? AddRange(IEnumerable<Paragraph> paragraphs);

    /// <summary>
    /// Add a new table to the end of the container
    /// </summary>
    /// <param name="table">Table to add</param>
    /// <returns>Table reference - may be copied if original table was already in document.</returns>
    Table Add(Table table);

    /// <summary>
    /// Insert a Table into this document.
    /// </summary>
    /// <param name="index">The index to insert this Table at.</param>
    /// <param name="table">The Table to insert.</param>
    /// <returns>The Table now associated with this document.</returns>
    Table Insert(int index, Table table);
}