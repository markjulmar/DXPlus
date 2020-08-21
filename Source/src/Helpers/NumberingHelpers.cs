using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    public static class NumberingHelpers
    {
        internal static int CreateNewNumberingSection(DocX document, ListItemType listType, int startNumber)
        {
            int numId = GetMaxNumId(document) + 1;
            int abstractNumId = GetMaxAbstractNumId(document) + 1;
            var listTemplate = listType switch
            {
                ListItemType.Bulleted => Resources.DefaultBulletNumberingXml(HelperFunctions.GenerateHexId()),
                ListItemType.Numbered => Resources.DefaultDecimalNumberingXml(HelperFunctions.GenerateHexId()),
                _ => throw new InvalidOperationException($"Unable to deal with ListItemType: {listType}."),
            };

            var abstractNumTemplate = listTemplate.FirstLocalNameDescendant("abstractNum");
            abstractNumTemplate.SetAttributeValue(DocxNamespace.Main + "abstractNumId", abstractNumId);

            var abstractNumIdElement = new XElement(DocxNamespace.Main + "abstractNumId", new XAttribute(DocxNamespace.Main + "val", abstractNumId));
            var abstractNumXml = new XElement(DocxNamespace.Main + "num", new XAttribute(DocxNamespace.Main + "numId", numId), abstractNumIdElement);

            if (startNumber != 1)
            {
                var startOverride = new XElement(DocxNamespace.Main + "lvlOverride",
                    new XAttribute(DocxNamespace.Main + "ilvl", 0),
                    new XElement(DocxNamespace.Main + "startOverride",
                        new XAttribute(DocxNamespace.Main + "val", startNumber)));
                abstractNumXml.Add(startOverride);
            }

            var abstractNumNode = document.numberingDoc.Root!.Descendants().LastOrDefault(xElement => xElement.Name.LocalName == "abstractNum");
            var numXml = document.numberingDoc.Root.Descendants().LastOrDefault(xElement => xElement.Name.LocalName == "num");

            if (abstractNumNode == null || numXml == null)
            {
                document.numberingDoc.Root.Add(abstractNumTemplate);
                document.numberingDoc.Root.Add(abstractNumXml);
            }
            else
            {
                abstractNumNode.AddAfterSelf(abstractNumTemplate);
                numXml.AddAfterSelf(abstractNumXml);
            }

            return numId;
        }

        internal static XElement GetNumElementFromNumId(DocX document, int numId)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (document.numberingDoc == null)
                throw new ArgumentException("Document missing numbering.xml");

            return document.numberingDoc
                .LocalNameDescendants("num")
                .FindByAttrVal(DocxNamespace.Main + "numId", numId.ToString());
        }

        internal static XElement GetAbstractNumFromNumId(DocX document, int numId)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (document.numberingDoc == null)
                throw new ArgumentException("Document missing numbering.xml");

            var numSection = document.numberingDoc
                .LocalNameDescendants("num")
                .Where(e => e.AttributeValue(DocxNamespace.Main + "numId") == numId.ToString())
                .Select(e => int.Parse(e.Element(DocxNamespace.Main + "abstractNumId").GetVal()))
                .Single();


            return document.numberingDoc.Root!
                .Descendants(DocxNamespace.Main + "abstractNum")
                .Single(e => e.AttributeValue(DocxNamespace.Main + "abstractNumId").Equals(numSection.ToString()));
        }

        /// <summary>
        /// Method to determine the last abstractNumId for a list element.
        /// Also useful for determining the next abstractNumId to use for inserting a new list element into the document.
        /// </summary>
        /// <returns>
        /// -1 if there are no elements in the list already.
        /// Increment the return for the next valid value of a new list element.
        /// </returns>
        internal static int GetMaxAbstractNumId(DocX document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (document.numberingDoc == null)
                throw new ArgumentException("Document missing numbering.xml");

            var absNums = document.numberingDoc.LocalNameDescendants("abstractNum").ToList();
            if (absNums.Count > 0)
            {
                return absNums.Attributes(DocxNamespace.Main + "abstractNumId")
                              .Max(e => int.Parse(e.Value));
            }

            return -1;
        }
        /// <summary>
        /// Method to determine the last numId for a list element.
        /// Also useful for determining the next numId to use for inserting a new list element into the document.
        /// </summary>
        /// <returns>
        /// 0 if there are no elements in the list already.
        /// Increment the return for the next valid value of a new list element.
        /// </returns>
        internal static int GetMaxNumId(DocX document)
        {
            if (document == null) 
                throw new ArgumentNullException(nameof(document));

            if (document.numberingDoc == null)
                throw new ArgumentException("Document missing numbering.xml");

            var numberSections = document.numberingDoc.LocalNameDescendants("num").ToList();
            if (numberSections.Count > 0)
            {
                return numberSections.Attributes(DocxNamespace.Main + "numId")
                                     .Max(e => int.Parse(e.Value));
            }

            return 0;
        }

        /// <summary>
        /// Returns the starting number (with override) for the given NumId and Level
        /// </summary>
        /// <param name="document">Document</param>
        /// <param name="numId">NumId</param>
        /// <param name="level">Level</param>
        /// <returns></returns>
        public static int GetStartingNumber(DocX document, int numId, int level = 0)
        {
            var numEl = GetNumElementFromNumId(document, numId);
            if (numEl == null)
                throw new ArgumentException($"NumId {numId} doesn't existing in numbering.xml", nameof(numId));

            // See if there's an override.
            var lvlOverride = numEl.Elements(DocxNamespace.Main + "lvlOverride")
                .FindByAttrVal(DocxNamespace.Main + "ilvl", level.ToString());
            if (lvlOverride != null)
            {
                return int.Parse(lvlOverride.Element(DocxNamespace.Main + "startOverride").GetVal());
            }

            // Otherwise, grab the abstract.
            var absNum = GetAbstractNumFromNumId(document, numId);
            if (absNum == null)
                throw new ArgumentException($"NumId {numId} [abstract] doesn't existing in numbering.xml", nameof(numId));

            var levelNode = absNum.Descendants(DocxNamespace.Main + "lvl")
                .FindByAttrVal(DocxNamespace.Main + "ilvl", level.ToString());
            return int.Parse(levelNode.Element(DocxNamespace.Main + "start").GetVal());
        }
    }

    internal static class ParagraphListHelpers
    {
        /// <summary>
        /// Determine if this paragraph is a list element.
        /// </summary>
        internal static bool IsListItem(this Paragraph p)
        {
            return ParagraphNumberProperties(p) != null;
        }

        /// <summary>
        /// Fetch the paragraph number properties for a list element.
        /// </summary>
        internal static XElement ParagraphNumberProperties(this Paragraph p)
        {
            return p.Xml.FirstLocalNameDescendant("numPr");
        }

        /// <summary>
        /// Return the associated numId for the list
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        internal static int GetListNumId(this Paragraph p)
        {
            var numProperties = ParagraphNumberProperties(p);
            return numProperties == null
                ? -1
                : int.Parse(numProperties.Element(DocxNamespace.Main + "numId").GetVal());
        }

        /// <summary>
        /// Return the list level for this paragraph
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        internal static int GetListLevel(this Paragraph p)
        {
            var numProperties = ParagraphNumberProperties(p);
            return numProperties == null
                ? -1
                : int.Parse(numProperties.Element(DocxNamespace.Main + "ilvl").GetVal());
        }

        /// <summary>
        /// Get the ListItemType property for the paragraph.
        /// Defaults to numbered if a list is found but the type is not specified
        /// </summary>
        internal static ListItemType GetListItemType(this Paragraph p)
        {
            var numProperties = ParagraphNumberProperties(p);
            if (numProperties == null)
                return ListItemType.None;

            string level = numProperties.Element(DocxNamespace.Main + "ilvl").GetVal();
            string numIdRef = numProperties.Element(DocxNamespace.Main + "numId").GetVal();

            // Find the number definition instance. We map <w:num> to <w:abstractNum>
            var numNode = p.Document.numberingDoc.LocalNameDescendants("num")?.FindByAttrVal(DocxNamespace.Main + "numId", numIdRef);
            if (numNode == null)
            {
                throw new Exception(
                    $"Number reference w:numId('{numIdRef}') used in document but not defined in numbering.xml");
            }

            // Get the abstractNumId
            string absNumId = numNode.FirstLocalNameDescendant("abstractNumId").GetVal();

            // Find the numbering style section
            var absNumNode = p.Document.numberingDoc.LocalNameDescendants("abstractNum")
                .FindByAttrVal(DocxNamespace.Main + "abstractNumId", absNumId);

            // Get the numbering format.
            var format = absNumNode.LocalNameDescendants("lvl")
                .FindByAttrVal(DocxNamespace.Main + "ilvl", level)
                .FirstLocalNameDescendant("numFmt");

            return format.TryGetEnumValue(out ListItemType result)
                ? result
                : ListItemType.Numbered;
        }
    }

}
