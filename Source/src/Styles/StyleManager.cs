using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Manager for the named styles (styles.xml) in the document.
    /// </summary>
    public sealed class StyleManager : DocXBase
    {
        private readonly XDocument stylesDoc;

        /// <summary>
        /// A list of all the available numbering styles in this document.
        /// </summary>
        public IEnumerable<Style> Styles =>
            Xml.Elements(Namespace.Main + "style").Select(e => new Style(e));

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="documentOwner">Owning document</param>
        /// <param name="stylesPart">Numbering part</param>
        public StyleManager(IDocument documentOwner, PackagePart stylesPart) : base(documentOwner, null)
        {
            PackagePart = stylesPart ?? throw new ArgumentNullException(nameof(stylesPart));
            stylesDoc = stylesPart.Load();
            Xml = stylesDoc.Root;
        }

        /// <summary>
        /// Save the changes back to the package.
        /// </summary>
        public void Save()
        {
            PackagePart.Save(stylesDoc);
        }

        /// <summary>
        /// Returns whether the given style exists in the style catalog
        /// </summary>
        /// <param name="styleId"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public bool HasStyle(string styleId, StyleType type)
        {
            return stylesDoc.Descendants(Namespace.Main + "style").Any(x =>
                x.AttributeValue(Namespace.Main + "type").Equals(type.GetEnumName())
                && x.AttributeValue(Namespace.Main + "styleId").Equals(styleId));
        }

        /// <summary>
        /// This method retrieves the XML block associated with a style.
        /// </summary>
        /// <param name="styleId">Id</param>
        /// <param name="type">Style type</param>
        /// <returns>Style if present</returns>
        internal Style GetStyle(string styleId, StyleType type) => 
            Styles.SingleOrDefault(s => s.Id == styleId && s.Type == type);

        /// <summary>
        /// This method adds a new Style XML block to the /word/styles.xml document
        /// </summary>
        /// <param name="xml">XML to add</param>
        internal void Add(XElement xml)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));

            if (xml.Name != Namespace.Main + "style")
                throw new ArgumentException("Passed element is not a <style> object.", nameof(xml));

            stylesDoc.Root!.Add(xml);
        }
    }
}
