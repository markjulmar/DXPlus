using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

namespace DXPlus
{
    public interface IContainer : IDocumentBase
    {
        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        ReadOnlyCollection<Paragraph> Paragraphs { get; }

        /// <summary>
        /// Returns all the sections associated with this container.
        /// </summary>
        List<Section> Sections { get; }

        /// <summary>
        /// Retrieve a list of all Table objects in the document
        /// </summary>
        IEnumerable<Table> Tables { get; }

        /// <summary>
        /// Retrieve a list of all Lists in the container (numbered or bulleted)
        /// </summary>
        IEnumerable<List> Lists { get; }

        /// <summary>
        /// Retrieve a list of all hyperlinks in the document
        /// </summary>
        IEnumerable<Hyperlink> Hyperlinks { get; }

        /// <summary>
        /// Retrieve a list of all images (pictures) in the document
        /// </summary>
        IEnumerable<Picture> Pictures { get; }

        /// <summary>
        /// Removes paragraph at specified position
        /// </summary>
        /// <param name="index">Index of paragraph to remove</param>
        /// <returns>True if removed</returns>
        bool RemoveParagraph(int index);

        /// <summary>
        /// Removes paragraph
        /// </summary>
        /// <param name="paragraph">Paragraph to remove</param>
        /// <returns>True if removed</returns>
        bool RemoveParagraph(Paragraph paragraph);

        /// <summary>
        /// Sets the Direction of content.
        /// </summary>
        /// <param name="direction">Direction either LeftToRight or RightToLeft</param>
        void SetDirection(Direction direction);

        /// <summary>
        /// Replace matched text with a new value
        /// </summary>
        /// <param name="searchValue">Text value to search for</param>
        /// <param name="newValue">Replacement value</param>
        /// <param name="options">Regex options</param>
        /// <param name="newFormatting">New formatting to apply</param>
        /// <param name="matchFormatting">Formatting to match</param>
        /// <param name="formattingOptions">Match formatting options</param>
        /// <param name="escapeRegEx">True to escape Regex expression</param>
        /// <param name="useRegExSubstitutions">True to use RegEx in substitution</param>
        void ReplaceText(string searchValue, string newValue,
            RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, 
            MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch, 
            bool escapeRegEx = true, bool useRegExSubstitutions = false);

        /// <summary>
        /// Insert a text block at a specific bookmark
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <param name="toInsert">Text to insert</param>
        bool InsertAtBookmark(string bookmarkName, string toInsert);

        /// <summary>
        /// Insert a paragraph into this container at a specific index
        /// </summary>
        /// <param name="index">Index to insert into</param>
        /// <param name="p">Paragraph to insert</param>
        /// <returns>Inserted paragraph</returns>
        Paragraph InsertParagraph(int index, Paragraph p);

        /// <summary>
        /// Add a paragraph at the end of the container
        /// </summary>
        Paragraph AddParagraph(Paragraph paragraph);

        /// <summary>
        /// Insert a new paragraph using the passed text.
        /// </summary>
        /// <param name="index">Index to insert into</param>
        /// <param name="text">Text for new paragraph</param>
        /// <param name="formatting">Formatting for new paragraph</param>
        /// <returns></returns>
        Paragraph InsertParagraph(int index, string text, Formatting formatting);

        /// <summary>
        /// Add a new section to the container
        /// </summary>
        void AddSection();

        /// <summary>
        /// Add a new page break to the container
        /// </summary>
        void AddPageBreak();

        /// <summary>
        /// Add a paragraph with the given text to the end of the container
        /// </summary>
        /// <param name="text">Text to add</param>
        /// <param name="formatting">Formatting to use</param>
        /// <returns></returns>
        Paragraph AddParagraph(string text, Formatting formatting);

        /// <summary>
        /// Add a new table to the end of the container
        /// </summary>
        /// <param name="table">Table to add</param>
        /// <returns>Table reference - may be copied if original table was already in document.</returns>
        Table AddTable(Table table);

        /// <summary>
        /// Insert a Table into this document. The Table's source can be a completely different document.
        /// </summary>
        /// <param name="table">The Table to insert.</param>
        /// <param name="index">The index to insert this Table at.</param>
        /// <returns>The Table now associated with this document.</returns>
        Table InsertTable(int index, Table table);

        /// <summary>
        /// Insert a List into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <returns>The List now associated with this document.</returns>
        List AddList(List list);

        /// <summary>
        /// Insert a List with a specific font size into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <param name="fontSize">Font size</param>
        /// <returns>The List now associated with this document.</returns>
        List AddList(List list, double fontSize);

        /// <summary>
        /// Insert a List with a specific font size/family into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <param name="fontFamily">Font family</param>
        /// <param name="fontSize">Font size</param>
        /// <returns>The List now associated with this document.</returns>
        List AddList(List list, System.Drawing.FontFamily fontFamily, double fontSize);

        /// <summary>
        /// Insert a List at a specific position.
        /// </summary>
        /// <param name="index">Position to insert into</param>
        /// <param name="list">The List to insert</param>
        /// <returns>The List now associated with this document.</returns>
        List InsertList(int index, List list);
    }
}