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
        Dictionary<string, string> CoreProperties { get; }

        /// <summary>
        /// Returns a list of custom properties in this document.
        /// </summary>
        Dictionary<string, CustomProperty> CustomProperties { get; }

        /// <summary>
        /// Returns true if any editing restrictions are imposed on this document.
        /// </summary>
        bool IsProtected { get; }

        /// <summary>
        /// Add a core property to this document. If a core property already exists with the same name it will be replaced. Core property names are case insensitive.
        /// </summary>
        ///<param name="name">The property name.</param>
        ///<param name="value">The property value.</param>
        void AddCoreProperty(string name, string value);

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        /// <param name="property">The CustomProperty to add to this document.</param>
        void AddCustomProperty(CustomProperty property);

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

        /// <summary>
        /// Add editing protection to this document.
        /// </summary>
        /// <param name="editRestrictions">The type of protection to add to this document.</param>
        void AddProtection(EditRestrictions editRestrictions);

        /// <summary>
        /// Add edit restrictions with a password
        /// </summary>
        /// <param name="editRestrictions"></param>
        /// <param name="password"></param>
        void AddProtection(EditRestrictions editRestrictions, string password);

        /// <summary>
        /// Returns the type of editing protection imposed on this document.
        /// </summary>
        /// <returns>The type of editing protection imposed on this document.</returns>
        EditRestrictions GetProtectionType();

        /// <summary>
        /// Remove editing protection from this document.
        /// </summary>
        void RemoveProtection();

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