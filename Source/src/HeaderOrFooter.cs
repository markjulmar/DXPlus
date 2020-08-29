using DXPlus.Helpers;
using System;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DXPlus
{
    public enum HeaderFooterType
    {
        First,
        Even,
        [XmlAttribute("default")]
        Odd,
    }

    /// <summary>
    /// Base class for header/footer
    /// </summary>
    public abstract class HeaderOrFooter : Container
    {
        /// <summary>
        /// The type of header/footer (even/odd/default)
        /// </summary>
        public HeaderFooterType Type { get; set; }

        /// <summary>
        /// True/False whether the header/footer has been created and exists.
        /// Setting this property will create/destroy the header/footer
        /// </summary>
        public bool Exists => Id != null && Xml != null;

        /// <summary>
        /// Relationship id for the header/footer
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Get the URI for this header/footer
        /// </summary>
        public Uri Uri => PackagePart?.Uri;

        /// <summary>
        /// PageNumber setting
        /// </summary>
        public abstract bool PageNumbers { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        internal HeaderOrFooter() : base(null, null)
        {
        }

        // Methods used to create/delete this header/footer.
        internal Action<HeaderOrFooter> CreateFunc;
        internal Action<HeaderOrFooter> DeleteFunc;

        /// <summary>
        /// Creates the header/footer and returns the generated paragraph.
        /// </summary>
        /// <returns></returns>
        public Paragraph Add()
        {
            if (!Exists)
                CreateFunc.Invoke(this);
            return Paragraphs[0];
        }

        /// <summary>
        /// Removes this header/footer
        /// </summary>
        public void Remove()
        {
            if (Exists)
            {
                DeleteFunc.Invoke(this);
                Xml = null;
                Id = null;
                PackagePart = null;
                Document = null;
            }
        }

        /// <summary>
        /// Save the header/footer out to disk.
        /// </summary>
        internal void Save()
        {
            if (Exists)
            {
                PackagePart.Save(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), Xml));
            }
        }
    }
}