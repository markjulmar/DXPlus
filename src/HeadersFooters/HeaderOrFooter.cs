using DXPlus.Helpers;
using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Base class for header/footer
    /// </summary>
    public abstract class HeaderOrFooter : BlockContainer
    {
        /// <summary>
        /// This is the actual Xml that gives this element substance.
        /// </summary>
        internal override XElement Xml
        {
            get
            {
                if (!Exists)
                    CreateFunc.Invoke(this);
                return base.Xml;
            }
            set => base.Xml = value;
        }

        /// <summary>
        /// The type of header/footer (even/odd/default)
        /// </summary>
        public HeaderFooterType Type { get; set; }

        /// <summary>
        /// True/False whether the header/footer has been created and exists.
        /// </summary>
        public bool Exists => Id != null && base.Xml != null && ExistsFunc(Id);

        /// <summary>
        /// Retrieves the first (main) paragraph for the header/footer.
        /// </summary>
        public Paragraph MainParagraph => Paragraphs.First();

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
        /// Called to add XML to this element.
        /// </summary>
        /// <param name="xml">XML to add</param>
        /// <returns>Element</returns>
        protected override XElement AddElementToContainer(XElement xml)
        {
            if (!Exists)
                CreateFunc.Invoke(this);

            return base.AddElementToContainer(xml);
        }

        /// <summary>
        /// Called to add new sections.
        /// </summary>
        /// <param name="breakType"></param>
        /// <returns></returns>
        public override Section AddSection(SectionBreakType breakType)
        {
            throw new InvalidOperationException("Cannot add sections to header/footer elements.");
        }

        /// <summary>
        /// Add a new page break to the container
        /// </summary>
        public override void AddPageBreak()
        {
            throw new InvalidOperationException("Cannot add page breaks to header/footer elements.");
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