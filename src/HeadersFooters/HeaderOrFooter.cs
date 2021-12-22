﻿using DXPlus.Helpers;
using System;
using System.Linq;
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
    public abstract class HeaderOrFooter : BlockContainer
    {
        /// <summary>
        /// The type of header/footer (even/odd/default)
        /// </summary>
        public HeaderFooterType Type { get; set; }

        /// <summary>
        /// True/False whether the header/footer has been created and exists.
        /// Setting this property will create/destroy the header/footer
        /// </summary>
        public bool Exists => Id != null && Xml != null && ExistsFunc(Id);

        /// <summary>
        /// Relationship id for the header/footer
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Get the URI for this header/footer
        /// </summary>
        public Uri Uri => PackagePart?.Uri;

        /// <summary>
        /// Constructor
        /// </summary>
        internal HeaderOrFooter() : base(null, null, null)
        {
        }

        // Methods used to create/delete this header/footer.
        internal Action<HeaderOrFooter> CreateFunc;
        internal Action<HeaderOrFooter> DeleteFunc;
        internal Func<string, bool> ExistsFunc;

        /// <summary>
        /// Creates the header/footer and returns the generated paragraph.
        /// </summary>
        /// <returns></returns>
        public Paragraph Add()
        {
            if (!Exists)
                CreateFunc.Invoke(this);
            return Paragraphs.First();
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
                SetOwner(null, null);
            }
        }

        /// <summary>
        /// Save the header/footer out to disk.
        /// </summary>
        internal void Save()
        {
            if (Exists)
            {
                PackagePart.Save(new XDocument(
                    new XDeclaration("1.0", "UTF-8", "yes"), Xml));
            }
        }
    }
}