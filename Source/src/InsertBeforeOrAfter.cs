using DXPlus.Helpers;
using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This class is the basis for Paragraphs, Lists, and Table elements. It provides helper methods to
    /// insert/add other elements before and after in the element tree.
    /// </summary>
    public abstract class InsertBeforeOrAfter : DocXElement
    {
        private Container owner;

        /// <summary>
        /// Default constructor used with Lists/Tables
        /// </summary>
        internal InsertBeforeOrAfter()
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document"></param>
        /// <param name="xml"></param>
        protected InsertBeforeOrAfter(IDocument document, XElement xml) : base(document, xml)
        {
        }

        /// <summary>
        /// Add a page break after the current element.
        /// </summary>
        public void AddPageBreak() => Xml.AddAfterSelf(PageBreak);

        /// <summary>
        /// Insert a page break before the current element.
        /// </summary>
        public void InsertPageBreakBefore() => Xml.AddBeforeSelf(PageBreak);

        /// <summary>
        /// Add a new paragraph after the current element.
        /// </summary>
        /// <param name="paragraph">Paragraph to insert</param>
        public void AddParagraph(Paragraph paragraph)
        {
            if (paragraph.Container != null)
                throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

            Xml.AddAfterSelf(paragraph.Xml);
            paragraph.Container = this.Container;
        }

        /// <summary>
        /// Add a paragraph after the current element using the passed text
        /// </summary>
        /// <param name="text">Text for new paragraph</param>
        /// <param name="formatting">Formatting for the paragraph</param>
        /// <returns>Newly created paragraph</returns>
        public Paragraph AddParagraph(string text, Formatting formatting)
        {
            XElement newParagraph = Paragraph.Create(text, formatting);
            Xml.AddAfterSelf(newParagraph);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();
            return new Paragraph(Document, newlyInserted, -1) {Container = Container};
        }

        /// <summary>
        /// Insert a paragraph before the current element
        /// </summary>
        /// <param name="paragraph"></param>
        public void InsertParagraphBefore(Paragraph paragraph)
        {
            if (paragraph.Container != null)
                throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

            Xml.AddBeforeSelf(paragraph.Xml);
            paragraph.Container = Container;
        }

        /// <summary>
        /// Insert a paragraph before the current element
        /// </summary>
        /// <param name="text">Text to use for new paragraph</param>
        /// <param name="formatting">Formatting to use</param>
        /// <returns></returns>
        public Paragraph InsertParagraphBefore(string text, Formatting formatting)
        {
            XElement paragraph = Paragraph.Create(text, formatting);
            Xml.AddBeforeSelf(paragraph);

            string id = paragraph.AttributeValue(Name.ParagraphId);
            return Container != null
                ? Container.Paragraphs.Single(p => p.Id == id)
                : new Paragraph(Document, paragraph, -1);
        }

        /// <summary>
        /// Add a new table after this container
        /// </summary>
        /// <param name="table">Table to add</param>
        public void AddTable(Table table)
        {
            if (table.Container != null)
                throw new ArgumentException("Cannot add table multiple times.", nameof(table));

            table.Container = this.Container;
            Xml.AddAfterSelf(table.Xml);
        }

        /// <summary>
        /// Insert a table before this element
        /// </summary>
        /// <param name="table"></param>
        public void InsertTableBefore(Table table)
        {
            if (table.Container != null)
                throw new ArgumentException("Cannot add table multiple times.", nameof(table));

            table.Container = Container;
            Xml.AddBeforeSelf(table.Xml);
        }

        /// <summary>
        /// Set the container owner for this element.
        /// </summary>
        internal Container Container
        {
            get => owner;
            set
            {
                if (owner == value)
                    return;

                if (owner != null)
                    OnRemovedFromContainer(owner);

                owner = value;
                PackagePart = owner?.PackagePart;
                Document = owner?.Document;

                if (owner != null)
                    OnAddedToContainer(owner);
            }
        }

        /// <summary>
        /// Invoked when this element is added to a container.
        /// </summary>
        /// <param name="container">Current Container owner</param>
        protected virtual void OnAddedToContainer(Container container)
        {
        }

        /// <summary>
        /// Invoked when this element is removed from a container.
        /// </summary>
        /// <param name="container">Current Container owner</param>
        protected virtual void OnRemovedFromContainer(Container container)
        {
        }

        /// <summary>
        /// Page break element
        /// </summary>
        private static XElement PageBreak => new XElement(Name.Paragraph,
            new XAttribute(Name.ParagraphId, HelperFunctions.GenerateHexId()),
            new XElement(Name.Run,
                new XElement(Namespace.Main + "br",
                    new XAttribute(Namespace.Main + "type", "page"))));
    }
}