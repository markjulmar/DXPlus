using DXPlus.Charts;
using System;
using System.Collections.Generic;
using System.IO;

namespace DXPlus
{
    /// <summary>
    /// This represents a single document
    /// </summary>
    public interface IDocument : IContainer, IDisposable
    {
        /// <summary>
        /// Editing session id for this session.
        /// </summary>
        string RevisionId { get; }

        /// <summary>
        /// Numbering styles in this document. Null if no lists are in use.
        /// </summary>
        NumberingStyleManager NumberingStyles { get; }

        /// <summary>
        /// The style manager for this document. Null if no styles are present.
        /// </summary>
        StyleManager Styles { get; }

        /// <summary>
        /// Retrieve all bookmarks
        /// </summary>
        BookmarkCollection Bookmarks { get; }

        /// <summary>
        /// Should the Document use different Headers and Footers for odd and even pages?
        /// </summary>
        bool DifferentEvenOddHeadersFooters { get; set; }

        /// <summary>
        /// Get the text of each endnote from this document
        /// </summary>
        IEnumerable<string> EndnotesText { get; }

        /// <summary>
        /// Get the text of each footnote from this document
        /// </summary>
        IEnumerable<string> FootnotesText { get; }

        /// <summary>
        /// Returns a list of Images in this document.
        /// </summary>
        List<Image> Images { get; }

        ///<summary>
        /// Returns the list of document core properties with corresponding values.
        ///</summary>
        IReadOnlyDictionary<DocumentPropertyName, string> DocumentProperties { get; }

        /// <summary>
        /// Returns a list of custom properties in this document.
        /// </summary>
        IReadOnlyDictionary<string, object> CustomProperties { get; }

        /// <summary>
        /// Set a known core property value
        /// </summary>
        /// <param name="name">Property to set</param>
        /// <param name="value">Value to assign</param>
        void SetPropertyValue(DocumentPropertyName name, string value);

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        void AddCustomProperty(string name, string value);

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        void AddCustomProperty(string name, double value);

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        void AddCustomProperty(string name, bool value);

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        void AddCustomProperty(string name, DateTime value);

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        void AddCustomProperty(string name, int value);

        /// <summary>
        /// Add an Image into this document from a fully qualified or relative filename.
        /// </summary>
        /// <param name="imageFileName">The fully qualified or relative filename.</param>
        /// <returns>An Image file.</returns>
        Image AddImage(string imageFileName);

        /// <summary>
        /// Add an Image into this document from a Stream.
        /// </summary>
        /// <param name="imageStream">A Stream stream.</param>
        /// <param name="contentType">Content type - image/jpg</param>
        /// <returns>An Image file.</returns>
        Image AddImage(Stream imageStream, string contentType = "image/jpg");

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateFilePath">The path to the document template file.</param>
        ///<exception cref="FileNotFoundException">The document template file not found.</exception>
        void ApplyTemplate(string templateFilePath);

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateFilePath">The path to the document template file.</param>
        ///<param name="includeContent">Whether to copy the document template text content to document.</param>
        ///<exception cref="FileNotFoundException">The document template file not found.</exception>
        void ApplyTemplate(string templateFilePath, bool includeContent);

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateStream">The stream of the document template file.</param>
        void ApplyTemplate(Stream templateStream);

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateStream">The stream of the document template file.</param>
        ///<param name="includeContent">Whether to copy the document template text content to document.</param>
        void ApplyTemplate(Stream templateStream, bool includeContent);

        /// <summary>
        /// Insert a chart in document
        /// </summary>
        void InsertChart(Chart chart);

        /// <summary>
        /// Inserts a default TOC into the current document.
        /// Title: Table of contents
        /// Switches will be: TOC \h \o '1-3' \u \z
        /// </summary>
        /// <returns>The inserted TableOfContents</returns>
        TableOfContents InsertDefaultTableOfContents();

        /// <summary>
        /// Inserts a TOC into the current document.
        /// </summary>
        /// <param name="title">The title of the TOC</param>
        /// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
        /// <param name="headerStyle">Lets you set the style name of the TOC header</param>
        /// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
        /// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
        /// <returns>The inserted TableOfContents</returns>
        TableOfContents InsertTableOfContents(string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null);

        /// <summary>
        /// Inserts at TOC into the current document before the provided <paramref name="reference"/>
        /// </summary>
        /// <param name="reference">The paragraph to use as reference</param>
        /// <param name="title">The title of the TOC</param>
        /// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
        /// <param name="headerStyle">Lets you set the style name of the TOC header</param>
        /// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
        /// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
        /// <returns>The inserted TableOfContents</returns>
        TableOfContents InsertTableOfContents(Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null);

        /// <summary>
        /// Close the document and release all resources.
        /// </summary>
        void Close();

        /// <summary>
        /// Save this document back to the location it was loaded from.
        /// </summary>
        void Save();

        /// <summary>
        /// Save this document to a file.
        /// </summary>
        /// <param name="newFileName">The filename to save this document as.</param>
        void SaveAs(string newFileName);

        /// <summary>
        /// Save this document to a Stream.
        /// </summary>
        /// <param name="newStreamDestination">The Stream to save this document to.</param>
        void SaveAs(Stream newStreamDestination);
    }
}