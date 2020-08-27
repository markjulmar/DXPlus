using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This class provides functions for inserting new DocXElements before or after the current DocXElement.
    /// </summary>
    public abstract class InsertBeforeOrAfter : DocXBase
    {
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

            if (this is Paragraph owner)
            {
                return new Paragraph(Document, newlyInserted, owner.EndIndex);
            }

            paragraph.Xml = newlyInserted;
            return paragraph;
        }

        /// <summary>
        /// Add a paragraph after the current element using the passed text
        /// </summary>
        /// <param name="text">Text for new paragraph</param>
        /// <param name="trackChanges">True to track changes</param>
        /// <param name="formatting">Formatting for the paragraph</param>
        /// <returns>Newly created paragraph</returns>
        public Paragraph AddParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
        {
            var newParagraph = ParagraphHelpers.Create(text, trackChanges, formatting);
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
        /// <param name="trackChanges">True to track changes</param>
        /// <param name="formatting">Formatting to use</param>
        /// <returns></returns>
        public Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
        {
            var newParagraph = ParagraphHelpers.Create(text, trackChanges, formatting);
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
            return new Table(Document, newlyInserted)
            {
                PackagePart = PackagePart,
                Design = table.Design
            };
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
            return new Table(Document, newlyInserted)
            {
                PackagePart = PackagePart,
                Design = table.Design
            };
        }

        private static XElement PageBreak => new XElement(Name.Paragraph,
            new XElement(Name.Run,
                new XElement(Namespace.Main + "br",
                    new XAttribute(Namespace.Main + "type", "page"))));
    }
}