using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a List in a document.
    /// </summary>
    public class List : InsertBeforeOrAfter
    {
        internal List(DocX document, XElement xml) : base(document, xml)
        {
        }

        /// <summary>
        /// This is a list of paragraphs that will be added to the document
        /// when the list is inserted into the document.
        /// The paragraph needs a numPr defined to be in this items collection.
        /// </summary>
        public List<Paragraph> Items { get; } = new List<Paragraph>();

        /// <summary>
        /// The ListItemType (bullet or numbered) of the list.
        /// </summary>
        public ListItemType? ListType { get; private set; }

        /// <summary>
        /// The numId used to reference the list settings in the numbering.xml
        /// </summary>
        public int NumId { get; private set; }

        /// <summary>
        /// Add an item to the list
        /// </summary>
        /// <param name="listText">Text</param>
        /// <param name="level">Level</param>
        /// <param name="listType">Type</param>
        /// <param name="startNumber">Starting number</param>
        /// <param name="trackChanges">True to track changes</param>
        /// <param name="continueNumbering">True to continue numbering</param>
        /// <returns></returns>
        public List AddItem(string listText, int level = 0, ListItemType listType = ListItemType.Numbered, 
            int? startNumber = null, bool trackChanges = false, bool continueNumbering = false)
        {
            if (startNumber.HasValue && continueNumbering) 
                throw new InvalidOperationException("Cannot specify a start number and at the same time continue numbering from another list");
            
            var listToReturn = HelperFunctions.CreateItemInList(this, listText, level, listType, startNumber, trackChanges, continueNumbering);
            var lastItem = listToReturn.Items.LastOrDefault();
            if (lastItem != null)
            {
                lastItem.packagePart = packagePart;
            }
            
            return listToReturn;
        }

        /// <summary>
        /// Adds an item to the list.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <exception cref="InvalidOperationException">
        /// Throws an InvalidOperationException if the item cannot be added to the list.
        /// </exception>
        public void AddItem(Paragraph paragraph)
        {
            if (paragraph.IsListItem)
            {
                var numIdNode = paragraph.Xml.Descendants().First(s => s.Name.LocalName == "numId");
                var numId = int.Parse(numIdNode.Attribute(DocxNamespace.Main + "val").Value);

                if (CanAddListItem(paragraph))
                {
                    NumId = numId;
                    Items.Add(paragraph);
                }
                else
                {
                    throw new InvalidOperationException("New list items can only be added to this list if they are have the same numId.");
                }
            }
        }

        public void AddItem(Paragraph paragraph, int start)
        {
            UpdateNumberingForLevelStartNumber(int.Parse(paragraph.IndentLevel.ToString()), start);

            if (ContainsLevel(start))
            {
                throw new InvalidOperationException("Cannot add a paragraph with a start value if another element already exists in this list with that level.");
            }

            AddItem(paragraph);
        }

        /// <summary>
        /// Determine if it is able to add the item to the list
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>
        /// Return true if AddItem(...) will succeed with the given paragraph.
        /// </returns>
        public bool CanAddListItem(Paragraph paragraph)
        {
            if (paragraph.IsListItem)
            {
                var numIdNode = paragraph.Xml.Descendants().First(s => s.Name.LocalName == "numId");
                var numId = int.Parse(numIdNode.Attribute(DocxNamespace.Main + "val").Value);
                if (NumId == 0 || (numId == NumId && numId > 0))
                {
                    return true;
                }
            }
            return false;
        }

        public bool ContainsLevel(int level) => Items.Any(i => i.ParagraphNumberProperties.FirstLocalNameDescendant("ilvl").Value == level.ToString());

        internal void CreateNewNumberingNumId(ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool continueNumbering = false)
        {
            ValidateDocXNumberingPartExists();
            if (Document.numbering.Root == null)
            {
                throw new InvalidOperationException("Numbering section did not instantiate properly.");
            }

            ListType = listType;

            var numId = GetMaxNumId() + 1;
            var abstractNumId = GetMaxAbstractNumId() + 1;
            XDocument listTemplate = listType switch
            {
                ListItemType.Bulleted => Resources.DefaultBulletNumberingXml,
                ListItemType.Numbered => Resources.DefaultDecimalNumberingXml,
                _ => throw new InvalidOperationException($"Unable to deal with ListItemType: {listType}."),
            };
            var abstractNumTemplate = listTemplate.Descendants().Single(d => d.Name.LocalName == "abstractNum");
            abstractNumTemplate.SetAttributeValue(DocxNamespace.Main + "abstractNumId", abstractNumId);

            var abstractNumXml = GetAbstractNumXml(abstractNumId, numId, startNumber, continueNumbering);
            var abstractNumNode = Document.numbering.Root.Descendants().LastOrDefault(xElement => xElement.Name.LocalName == "abstractNum");
            var numXml = Document.numbering.Root.Descendants().LastOrDefault(xElement => xElement.Name.LocalName == "num");

            if (abstractNumNode == null || numXml == null)
            {
                Document.numbering.Root.Add(abstractNumTemplate);
                Document.numbering.Root.Add(abstractNumXml);
            }
            else
            {
                abstractNumNode.AddAfterSelf(abstractNumTemplate);
                numXml.AddAfterSelf(abstractNumXml);
            }

            NumId = numId;
        }

        /// <summary>
        /// Get the abstractNum definition for the given numId
        /// </summary>
        /// <param name="numId">The numId on the pPr element</param>
        /// <returns>XElement representing the requested abstractNum</returns>
        internal XElement GetAbstractNum(int numId)
        {
            var num = Document.numbering.Descendants().First(d => d.Name.LocalName == "num" && d.AttributeValue(DocxNamespace.Main + "numId").Equals(numId.ToString()));
            var abstractNumId = num.Descendants().First(d => d.Name.LocalName == "abstractNumId");
            return Document.numbering.Descendants().First(d => d.Name.LocalName == "abstractNum" && d.AttributeValue("abstractNumId").Equals(abstractNumId.Value));
        }

        private XElement GetAbstractNumXml(int abstractNumId, int numId, int? startNumber, bool continueNumbering)
        {
            var startOverride = new XElement(DocxNamespace.Main + "startOverride", new XAttribute(DocxNamespace.Main + "val", startNumber ?? 1));
            var lvlOverride = new XElement(DocxNamespace.Main + "lvlOverride", new XAttribute(DocxNamespace.Main + "ilvl", 0), startOverride);
            var abstractNumIdElement = new XElement(DocxNamespace.Main + "abstractNumId", new XAttribute(DocxNamespace.Main + "val", abstractNumId));
            return continueNumbering
                ? new XElement(DocxNamespace.Main + "num", new XAttribute(DocxNamespace.Main + "numId", numId), abstractNumIdElement)
                : new XElement(DocxNamespace.Main + "num", new XAttribute(DocxNamespace.Main + "numId", numId), abstractNumIdElement, lvlOverride);
        }

        /// <summary>
        /// Method to determine the last abstractNumId for a list element.
        /// Also useful for determining the next abstractNumId to use for inserting a new list element into the document.
        /// </summary>
        /// <returns>
        /// -1 if there are no elements in the list already.
        /// Increment the return for the next valid value of a new list element.
        /// </returns>
        private int GetMaxAbstractNumId()
        {
            const int defaultValue = -1;

            if (Document.numbering == null)
                return defaultValue;

            var numlist = Document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum").ToList();
            if (numlist.Count > 0)
                return numlist.Attributes(DocxNamespace.Main + "abstractNumId").Max(e => int.Parse(e.Value));

            return defaultValue;
        }

        /// <summary>
        /// Method to determine the last numId for a list element.
        /// Also useful for determining the next numId to use for inserting a new list element into the document.
        /// </summary>
        /// <returns>
        /// 0 if there are no elements in the list already.
        /// Increment the return for the next valid value of a new list element.
        /// </returns>
        private int GetMaxNumId()
        {
            const int defaultValue = 0;
            if (Document.numbering == null)
                return defaultValue;

            var numlist = Document.numbering.Descendants().Where(d => d.Name.LocalName == "num").ToList();
            if (numlist.Count > 0)
                return numlist.Attributes(DocxNamespace.Main + "numId").Max(e => int.Parse(e.Value));
            return defaultValue;
        }

        private void UpdateNumberingForLevelStartNumber(int iLevel, int start)
        {
            var abstractNum = GetAbstractNum(NumId);
            var level = abstractNum.Descendants().First(el => el.Name.LocalName == "lvl" && el.AttributeValue(DocxNamespace.Main + "ilvl") == iLevel.ToString());
            level.Descendants().First(el => el.Name.LocalName == "start").SetAttributeValue(DocxNamespace.Main + "val", start);
        }

        private void ValidateDocXNumberingPartExists()
        {
            var numberingUri = new Uri("/word/numbering.xml", UriKind.Relative);

            // If the internal document contains no /word/numbering.xml create one.
            if (!Document.package.PartExists(numberingUri))
                Document.numbering = HelperFunctions.AddDefaultNumberingXml(Document.package);
        }
    }
}