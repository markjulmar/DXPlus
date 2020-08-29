using System.Buffers;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This class is the basis for Paragraphs, Lists, and Table elements. It provides helper methods to
    /// insert/add other elements before and after in the element tree.
    /// </summary>
    public abstract class InsertBeforeOrAfter : DocXBase
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
        public void AddPageBreakAfterSelf() => Xml.AddAfterSelf(PageBreak);

        /// <summary>
        /// Insert a page break before the current element.
        /// </summary>
        public void InsertPageBreakBeforeSelf() => Xml.AddBeforeSelf(PageBreak);

        /// <summary>
        /// Add a new paragraph after the current element.
        /// </summary>
        /// <param name="paragraph">Paragraph to insert</param>
        public Paragraph AddParagraphAfterSelf(Paragraph paragraph)
        {
            Xml.AddAfterSelf(paragraph.Xml);
            var newlyInserted = Xml.ElementsAfterSelf().First();

            if (this is Paragraph me)
            {
                return new Paragraph(Document, newlyInserted, me.EndIndex);
            }

            paragraph.Xml = newlyInserted;
            return paragraph;
        }

        /// <summary>
        /// Add a paragraph after the current element using the passed text
        /// </summary>
        /// <param name="text">Text for new paragraph</param>
        /// <param name="formatting">Formatting for the paragraph</param>
        /// <returns>Newly created paragraph</returns>
        public Paragraph AddParagraphAfterSelf(string text, Formatting formatting)
        {
            var newParagraph = ParagraphHelpers.Create(text, formatting);
            Xml.AddAfterSelf(newParagraph);
            var newlyInserted = Xml.ElementsAfterSelf().First();
            
            return new Paragraph(Document, newlyInserted, -1);
        }

        /// <summary>
        /// Insert a paragraph before the current element
        /// </summary>
        /// <param name="paragraph"></param>
        public Paragraph InsertParagraphBeforeSelf(Paragraph paragraph)
        {
            Xml.AddBeforeSelf(paragraph.Xml);
            paragraph.Xml = Xml.ElementsBeforeSelf().First();
            return paragraph;
        }

        /// <summary>
        /// Insert a paragraph before the current element
        /// </summary>
        /// <param name="text">Text to use for new paragraph</param>
        /// <param name="formatting">Formatting to use</param>
        /// <returns></returns>
        public Paragraph InsertParagraphBeforeSelf(string text, Formatting formatting)
        {
            var newParagraph = ParagraphHelpers.Create(text, formatting);
            Xml.AddBeforeSelf(newParagraph);
            var newlyInserted = Xml.ElementsBeforeSelf().Last();
            
            return new Paragraph(Document, newlyInserted, -1);
        }

        /// <summary>
        /// Add a new table after this container
        /// </summary>
        /// <param name="table">Table to add</param>
        public Table AddTableAfterSelf(Table table)
        {
            if (table.Document == null)
            {
                Xml.AddAfterSelf(table.Xml);
                table.Document = Document;
                return table;
            }

            Xml.AddAfterSelf(table.Xml);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();

            // Already owned by another document -- clone it.
            return new Table(Document, newlyInserted) { Design = table.Design };
        }

        /// <summary>
        /// Insert a table before this element
        /// </summary>
        /// <param name="table"></param>
        public Table InsertTableBeforeSelf(Table table)
        {
            if (table.Document == null)
            {
                Xml.AddBeforeSelf(table.Xml);
                table.Document = Document;
                return table;
            }

            Xml.AddBeforeSelf(table.Xml);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

            // Already owned by another document -- clone it.
            return new Table(Document, newlyInserted) { Design = table.Design };
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
            new XElement(Name.Run,
                new XElement(Namespace.Main + "br",
                    new XAttribute(Namespace.Main + "type", "page"))));
    }
}