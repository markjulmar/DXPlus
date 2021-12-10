using DXPlus.Helpers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This represents a single item in the list.
    /// </summary>
    public class ListItem
    {
        /// <summary>
        /// Internal constructor
        /// </summary>
        internal ListItem()
        {
        }

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
                return numIdVal != null && int.TryParse(numIdVal, out var result) ? result : 0;
            }
            set => Paragraph.ParagraphNumberProperties()
                    .Element(Namespace.Main + "numId")?
                    .SetAttributeValue(Name.MainVal, value);
        }

        public string Text => Paragraph.Text;
        internal XElement Xml => Paragraph.Xml;
    }

    /// <summary>
    /// Represents a List in a container.
    /// </summary>
    public class List : Block
    {
        private readonly List<ListItem> items = new();
        private FontFamily fontFamily;
        private double? fontSize;

        /// <summary>
        /// List of items to add - these will create paragraph objects in the document
        /// tied to numberId elements in numbering.xml
        /// </summary>
        public IReadOnlyList<ListItem> Items => items.AsReadOnly();

        /// <summary>
        /// The ListItemType (bullet or numbered) of the list.
        /// </summary>
        public NumberingFormat ListType { get; internal set; }

        /// <summary>
        /// Start number
        /// </summary>
        public int StartNumber { get; set; }

        /// <summary>
        /// The numId used to reference the list settings in the numbering.xml
        /// </summary>
        public int NumId { get; internal set; }

        /// <summary>
        /// Default font family to use
        /// </summary>
        public FontFamily Font
        {
            get => fontFamily;
            set
            {
                fontFamily = value;
                foreach (var item in Items)
                {
                    item.Paragraph.AddFormatting(new Formatting {Font = value});
                }
            }
        }

        /// <summary>
        /// Default font size to use
        /// </summary>
        public double? FontSize
        {
            get => fontSize;
            set
            {
                fontSize = value;
                foreach (var item in Items)
                {
                    item.Paragraph.AddFormatting(new Formatting { FontSize = value });
                }
            }
        }

        /// <summary>
        /// Internal constructor when building List from paragraphs
        /// </summary>
        internal List()
        {
            this.BlockContainer = new NullContainer();
        }

        /// <summary>
        /// Public constructor
        /// </summary>
        /// <param name="listType">List type</param>
        /// <param name="startNumber">Starting number</param>
        public List(NumberingFormat listType, int startNumber = 1) : this()
        {
            if (listType == NumberingFormat.None)
                throw new ArgumentException("Cannot use None as the ListType.", nameof(listType));

            ListType = listType;
            StartNumber = startNumber;
            NumId = 0;
        }

        /// <summary>
        /// Constructor to load a set of items into the list.
        /// </summary>
        /// <param name="listType">List type</param>
        /// <param name="items">Items to add</param>
        /// <param name="startNumber">Starting number</param>
        public List(NumberingFormat listType, IEnumerable<string> items, int startNumber = 1) : this(listType, startNumber)
        {
            foreach (var text in items)
            {
                AddItem(text);
            }
        }

        /// <summary>
        /// Add an item to the list
        /// </summary>
        /// <param name="text">Text</param>
        /// <param name="level">Level</param>
        /// <returns></returns>
        public List AddItem(string text, int level = 0)
        {
            if (text == null)
                throw new ArgumentNullException(nameof(text));

            var newParagraphSection = new XElement(Name.Paragraph,
                new XAttribute(Name.ParagraphId, HelperFunctions.GenerateHexId()),
                new XElement(Name.ParagraphProperties,
                    new XElement(Name.ParagraphStyle, new XAttribute(Name.MainVal, "ListParagraph")),
                    new XElement(Namespace.Main + "numPr",
                        new XElement(Namespace.Main + "ilvl", new XAttribute(Name.MainVal, level)),
                        new XElement(Namespace.Main + "numId", new XAttribute(Name.MainVal, NumId)))),
                new XElement(Name.Run, new XElement(Name.Text, text))
            );

            var newItem = new ListItem
            {
                Paragraph = new Paragraph(Document, newParagraphSection, 0, ContainerType.Paragraph)
                {
                    BlockContainer = this.BlockContainer,
                }
            };

            // Apply the font/size if it's set.
            if (Font != null || FontSize != null)
            {
                var formatting = new Formatting();
                if (Font != null) formatting.Font = Font;
                if (FontSize != null) formatting.FontSize = FontSize;
                newItem.Paragraph.AddFormatting(formatting);
            }

            BlockContainer.Xml.Add(newParagraphSection);

            items.Add(newItem);

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
                if (numIdVal != null && int.TryParse(numIdVal, out var result))
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

            paragraph.BlockContainer = BlockContainer;
            if (paragraph.DefaultFormatting == null)
            {
                paragraph.DefaultFormatting = new Formatting {Font = Font, FontSize = FontSize};
            }
            else
            {
                paragraph.DefaultFormatting.Font = Font;
                paragraph.DefaultFormatting.FontSize = FontSize;
            }

            BlockContainer.Xml.Add(paragraph);

            items.Add(new ListItem { Paragraph = paragraph });

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
                // Add the numbering.xml file
                doc.AddDefaultNumberingPart();

                // Create a numbering section if needed.
                if (NumId == 0)
                    NumId = doc.NumberingStyles.Create(ListType, StartNumber).Id;

                // Wire up the paragraphs
                foreach (var item in Items)
                {
                    item.Paragraph.Document = doc;
                    item.NumId = NumId;
                }
            }
        }

        /// <summary>
        /// Called when the package part is changed.
        /// </summary>
        protected override void OnPackagePartChanged(PackagePart previousValue, PackagePart newValue)
        {
            base.OnPackagePartChanged(previousValue, newValue);
            foreach (var item in Items)
                item.Paragraph.PackagePart = newValue;
        }
    }
}