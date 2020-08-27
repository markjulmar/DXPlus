using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// This represents a single item in the list.
    /// </summary>
    public class ListItem
    {
        /// <summary>
        /// Text
        /// </summary>
        public Paragraph Paragraph { get; set; }

        /// <summary>
        /// Indent level (0 == root)
        /// </summary>
        public int IndentLevel
        {
            get
            {
                string value = Paragraph?.ParagraphNumberProperties().FirstLocalNameDescendant("ilvl").GetVal();
                return value != null && int.TryParse(value, out int result) ? result : 0;
            }
        }

        /// <summary>
        /// Assigned NumId -- should match owner List
        /// </summary>
        public int NumId
        {
            get
            {
                string numIdVal = Paragraph.ParagraphNumberProperties().Element(Namespace.Main + "numId").GetVal();
                return numIdVal != null && int.TryParse(numIdVal, out int result) ? result : 0;
            }
            set => Paragraph.ParagraphNumberProperties()
                    .Element(Namespace.Main + "numId")?
                    .SetAttributeValue(Name.MainVal, value);
        }

        public string Text => Paragraph.Text;
        public XElement Xml => Paragraph.Xml;
        
    }

    /// <summary>
    /// Represents a List in a document.
    /// </summary>
    public class List : InsertBeforeOrAfter
    {
        /// <summary>
        /// List of items to add - these will create paragraph objects in the document
        /// tied to numberId elements in numbering.xml
        /// </summary>
        public List<ListItem> Items { get; } = new List<ListItem>();

        /// <summary>
        /// The ListItemType (bullet or numbered) of the list.
        /// </summary>
        public ListItemType ListType { get; internal set; }

        /// <summary>
        /// Start number
        /// </summary>
        public int StartNumber { get; set; }

        /// <summary>
        /// The numId used to reference the list settings in the numbering.xml
        /// </summary>
        public int NumId { get; private set; }

        /// <summary>
        /// Internal constructor when building List from paragraphs
        /// </summary>
        internal List()
        {
        }

        /// <summary>
        /// Public constructor
        /// </summary>
        public List(ListItemType listType, int startNumber = 1)
        {
            if (listType == ListItemType.None)
                throw new ArgumentException("Cannot use None as the ListType.", nameof(listType));

            ListType = listType;
            StartNumber = startNumber;
        }

        /// <summary>
        /// Method to clone a list
        /// </summary>
        /// <param name="otherList">Other list</param>
        /// <returns>Copy of the list</returns>
        internal List(List otherList)
        {
            ListType = otherList.ListType;
            StartNumber = otherList.StartNumber;
            this.Items.AddRange(otherList.Items.Select(item => new ListItem {Paragraph = new Paragraph {Xml = item.Xml}}));
        }

        /// <summary>
        /// Add an item to the list
        /// </summary>
        /// <param name="text">Text</param>
        /// <param name="level">Level</param>
        /// <param name="trackChanges">True to track changes</param>
        /// <returns></returns>
        public List AddItem(string text, int level = 0, bool trackChanges = false)
        {
            if (text == null) 
                throw new ArgumentNullException(nameof(text));

            var newParagraphSection = new XElement(Name.Paragraph,
                new XElement(Name.ParagraphProperties,
                    new XElement(Name.ParagraphStyle, new XAttribute(Name.MainVal, "ListParagraph")),
                    new XElement(Namespace.Main + "numPr",
                        new XElement(Namespace.Main + "ilvl", new XAttribute(Name.MainVal, level)),
                        new XElement(Namespace.Main + "numId", new XAttribute(Name.MainVal, 0)))),
                new XElement(Name.Run, new XElement(Name.Text, text))
            );

            if (trackChanges)
            {
                newParagraphSection = HelperFunctions.CreateEdit(EditType.Ins, DateTime.Now, newParagraphSection);
            }

            Items.Add(new ListItem { Paragraph = new Paragraph(null, newParagraphSection, 0, ContainerType.Paragraph) });

            return this;
        }

        /// <summary>
        /// Adds an item to the list.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="level"></param>
        public List AddItem(Paragraph paragraph, int level = 0)
        {
            var paraProps = paragraph.ParagraphNumberProperties();
            if (paraProps == null)
            {
                paraProps = new XElement(Name.ParagraphProperties,
                    new XElement(Name.ParagraphStyle, new XAttribute(Name.MainVal, "ListParagraph")),
                    new XElement(Namespace.Main + "numPr",
                        new XElement(Namespace.Main + "ilvl", new XAttribute(Name.MainVal, level)),
                        new XElement(Namespace.Main + "numId", new XAttribute(Name.MainVal, 0))));
                paragraph.Xml.AddFirst(paraProps);
            }
            else if (NumId == 0)
            {
                string numIdVal = paraProps.Element(Namespace.Main + "numId").GetVal();
                if (numIdVal != null && int.TryParse(numIdVal, out int result))
                {
                    if (Items.Any(i => i.NumId != 0 && i.NumId != result))
                    {
                        throw new InvalidOperationException(
                            "New list items can only be added to this list if they are have the same numId.");
                    }

                    NumId = result;
                }
            }

            if (!CanAddListItem(paragraph))
            {
                throw new InvalidOperationException(
                    "New list items can only be added to this list if they are have the same numId.");
            }

            Items.Add(new ListItem {Paragraph = paragraph});

            return this;
        }

        /// <summary>
        /// Determine if the given paragraph can be added to this list.
        ///    1. Paragraph has a w:numPr element.
        ///    2. w:numId == 0 or >0 and matches List.numId
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>
        /// Return true if paragraph can be added to the list.
        /// </returns>
        public bool CanAddListItem(Paragraph paragraph)
        {
            if (!paragraph.IsListItem()) 
                return false;
            
            var numIdNode = paragraph.Xml.FirstLocalNameDescendant("numId");
            if (numIdNode == null || !int.TryParse(numIdNode.GetVal(), out int numId))
                return false;

            return NumId == 0 || (numId == NumId && numId > 0);
        }

        protected override void OnDocumentOwnerChanged(IDocument previousValue, IDocument newValue)
        {
            base.OnDocumentOwnerChanged(previousValue, newValue);

            // Attached to a document!
            if (newValue is DocX doc)
            {
                PackagePart = doc.PackagePart;

                // Add the numbering.xml file
                doc.AddDefaultNumberingPart();

                // Create a numbering section if needed.
                if (NumId == 0)
                    NumId = NumberingHelpers.CreateNewNumberingSection(doc, ListType, StartNumber);

                // Wire up the paragraphs
                foreach (var item in Items)
                {
                    item.Paragraph.Document = doc;
                    item.NumId = NumId;
                }
            }
        }
    }
}