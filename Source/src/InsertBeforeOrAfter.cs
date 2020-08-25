using System;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// This class provides functions for inserting new DocXElements before or after the current DocXElement.
    /// Only certain DocXElements can support these functions without creating invalid documents, at the moment these are Paragraphs and Table.
    /// </summary>
    public abstract class InsertBeforeOrAfter : DocXElement
    {
        internal InsertBeforeOrAfter()
        {
        }

        internal InsertBeforeOrAfter(DocX document, XElement xml) : base(document, xml)
        {
        }

        private XElement PageBreak => new XElement(DocxNamespace.Main + "p",
                    new XElement(DocxNamespace.Main + "r",
                            new XElement(DocxNamespace.Main + "br",
                                new XAttribute(DocxNamespace.Main + "type", "page"))));

        public virtual void InsertPageBreakAfterSelf()
        {
            Xml.AddAfterSelf(PageBreak);
        }

        public virtual void InsertPageBreakBeforeSelf()
        {
            Xml.AddBeforeSelf(PageBreak);
        }

        public virtual Paragraph InsertParagraphAfterSelf(Paragraph p)
        {
            Xml.AddAfterSelf(p.Xml);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();

            if (this is Paragraph owner)
            {
                return new Paragraph(Document, newlyInserted, owner.EndIndex);
            }
            else
            {
                p.Xml = newlyInserted;
                return p;
            }
        }

        public virtual Paragraph InsertParagraphAfterSelf(string text)
        {
            return InsertParagraphAfterSelf(text, false, new Formatting());
        }

        public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
        {
            return InsertParagraphAfterSelf(text, trackChanges, new Formatting());
        }

        public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = CreateNewParagraph(text, trackChanges, formatting);
            Xml.AddAfterSelf(newParagraph);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();
            return new Paragraph(Document, newlyInserted, -1);
        }

        public virtual Paragraph InsertParagraphBeforeSelf(Paragraph p)
        {
            Xml.AddBeforeSelf(p.Xml);
            p.Xml = Xml.ElementsBeforeSelf().First();
            return p;
        }
        public virtual Paragraph InsertParagraphBeforeSelf(string text)
        {
            return InsertParagraphBeforeSelf(text, false, new Formatting());
        }

        public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
        {
            return InsertParagraphBeforeSelf(text, trackChanges, new Formatting());
        }

        public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = CreateNewParagraph(text, trackChanges, formatting);
            Xml.AddBeforeSelf(newParagraph);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();
            return new Paragraph(Document, newlyInserted, -1);
        }
        public virtual Table InsertTableAfterSelf(int rowCount, int columnCount)
        {
            XElement newTable = TableHelpers.CreateTable(rowCount, columnCount);
            Xml.AddAfterSelf(newTable);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();

            return new Table(Document, newlyInserted) { PackagePart = PackagePart };
        }

        public virtual Table InsertTableAfterSelf(Table t)
        {
            Xml.AddAfterSelf(t.Xml);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();

            return new Table(Document, newlyInserted) { PackagePart = PackagePart };
        }

        public virtual Table InsertTableBeforeSelf(int rowCount, int columnCount)
        {
            XElement newTable = TableHelpers.CreateTable(rowCount, columnCount);
            Xml.AddBeforeSelf(newTable);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

            return new Table(Document, newlyInserted) { PackagePart = PackagePart };
        }

        public virtual Table InsertTableBeforeSelf(Table t)
        {
            Xml.AddBeforeSelf(t.Xml);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();
            return new Table(Document, newlyInserted) { PackagePart = PackagePart }; //return new table, dont affect parameter table
        }

        private static XElement CreateNewParagraph(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = new XElement(DocxNamespace.Main + "p",
                new XElement(DocxNamespace.Main + "pPr"), HelperFunctions.FormatInput(text, formatting.Xml));

            if (trackChanges)
            {
                newParagraph = HelperFunctions.CreateEdit(EditType.Ins, DateTime.Now, newParagraph);
            }

            return newParagraph;
        }
    }
}