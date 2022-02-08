using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Helper methods to make a paragraph a ListItem.
    /// This involves adding a style + {w:numPr} element to indicate the numbering style.
    /// </summary>
    public static class ListExtensions
    {
        /// <summary>
        /// Determine if this paragraph is a list element.
        /// </summary>
        public static bool IsListItem(this Paragraph p)
            => p.ParagraphNumberProperties() != null
               || p.Properties?.StyleName == "ListParagraph";

        /// <summary>
        /// This ensures a secondary paragraph remains associated to a list.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static Paragraph ListStyle(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            // List items have a ListParagraph style.
            var paraProps = paragraph.Xml.GetOrInsertElement(Name.ParagraphProperties);
            var style = paraProps.GetOrAddElement(Name.ParagraphStyle);
            style.SetAttributeValue(Name.MainVal, "ListParagraph");

            return paragraph;
        }

        /// <summary>
        /// Make the passed paragraph part of a List (bullet or numbered).
        /// </summary>
        /// <param name="paragraph">Paragraph</param>
        /// <param name="numberingDefinition">Numbering style to use</param>
        /// <param name="level">Indent level (0-based, defaults to top level)</param>
        /// <returns>Paragraph</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static Paragraph ListStyle(this Paragraph paragraph, 
            NumberingDefinition numberingDefinition, int level = 0)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));
            if (numberingDefinition == null)
                throw new ArgumentNullException(nameof(numberingDefinition));
            if (level < 0)
                throw new ArgumentOutOfRangeException(nameof(level));

            // List items have a ListParagraph style.
            var paraProps = paragraph.Xml.GetOrInsertElement(Name.ParagraphProperties);
            var style = paraProps.GetOrAddElement(Name.ParagraphStyle);
            style.SetAttributeValue(Name.MainVal, "ListParagraph");

            // Add the paragraph numbering properties.
            var pnpElement = paragraph.ParagraphNumberProperties();
            if (pnpElement == null)
            {
                paraProps.Add(new XElement(Namespace.Main + "numPr",
                    new XElement(Namespace.Main + "ilvl", new XAttribute(Name.MainVal, level)),
                    new XElement(Namespace.Main + "numId", new XAttribute(Name.MainVal, numberingDefinition.Id))));
            }
            // Has list info? Replace it.
            else
            {
                pnpElement.GetOrAddElement(Namespace.Main + "ilvl").SetAttributeValue(Name.MainVal, level);
                pnpElement.GetOrAddElement(Namespace.Main + "numId")
                    .SetAttributeValue(Name.MainVal, numberingDefinition.Id);
            }

            return paragraph;
        }

        /// <summary>
        /// Fetch the paragraph number properties for a list element.
        /// </summary>
        private static XElement ParagraphNumberProperties(this Paragraph p)
            => p.Xml.FirstLocalNameDescendant("numPr");

        /// <summary>
        /// Return the associated numId for the list
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        internal static int? GetListNumberingDefinitionId(this Paragraph p)
        {
            if (!p.IsListItem())
                return null;

            var numProperties = ParagraphNumberProperties(p);
            if (numProperties == null)
            {
                // Backup and try to find a previous ListItem style with properties.
                // This paragraph would inherit that.
                p = p.PreviousParagagraph;
                while (p != null && p.IsListItem())
                {
                    numProperties = ParagraphNumberProperties(p);
                    if (numProperties != null)
                        break;
                    p = p.PreviousParagagraph;
                }
            }

            return numProperties == null
                ? null
                : int.Parse(numProperties.Element(Namespace.Main + "numId").GetVal());
        }

        /// <summary>
        /// Retrieve the associated numbering definition (if any) for the specified paragraph.
        /// </summary>
        /// <param name="p">Paragraph</param>
        /// <returns>Numbering definition, or null if not in a list.</returns>
        public static NumberingDefinition GetListNumberingDefinition(this Paragraph p)
        {
            if (!p.IsListItem()) 
                return null;

            var id = p.GetListNumberingDefinitionId();
            if (id == null) return null;

            var styles = p.Document.NumberingStyles;
            var definition = styles.Definitions.SingleOrDefault(d => d.Id == id.Value);
            if (definition == null)
            {
                throw new Exception(
                    $"Number reference w:numId('{id.Value}') used in document but not defined in /word/numbering.xml");
            }

            return definition;
        }

        /// <summary>
        /// Retrieves all the paragraphs for a list based on the numId.
        /// </summary>
        /// <param name="container">Container to apply to</param>
        /// <param name="numberingDefinitionId">Numbering definition id to group paragraphs by.</param>
        /// <returns>Paragraphs associated with this list</returns>
        public static IEnumerable<Paragraph> GetListById(this IContainer container, int numberingDefinitionId)
            => container.Paragraphs.Where(p => p.GetListNumberingDefinitionId() == numberingDefinitionId);

        /// <summary>
        /// Returns the index # of this paragraph if it's part of a list.
        /// </summary>
        /// <param name="p">Paragraph</param>
        /// <returns>Index or null if not in a list.</returns>
        public static int? GetListIndex(this Paragraph p)
        {
            int? listId = p.GetListNumberingDefinitionId();
            if (listId == null)
                return null;

            IContainer container = p.Container ?? p.Document;
            var theList = container?.GetListById(listId.Value).ToList();
            return theList?.FindIndex(p2 => p2.Equals(p));
        }

        /// <summary>
        /// Return the list level for this paragraph
        /// </summary>
        /// <param name="p">Paragraph</param>
        /// <returns></returns>
        public static int? GetListLevel(this Paragraph p)
        {
            var numProperties = ParagraphNumberProperties(p);
            return numProperties == null
                ? null
                : int.Parse(numProperties.Element(Namespace.Main + "ilvl").GetVal());
        }

        /// <summary>
        /// Get the ListItemType property for the paragraph.
        /// Defaults to numbered if a list is found but the type is not specified
        /// </summary>
        /// <param name="p">Paragraph</param>
        public static NumberingFormat GetNumberingFormat(this Paragraph p)
        {
            var numProperties = ParagraphNumberProperties(p);
            if (numProperties == null)
                return NumberingFormat.None;

            int numId = int.Parse(numProperties.Element(Namespace.Main + "numId").GetVal());

            // A value of 0 for the @val attribute indicates the removal of numbering properties at
            // a particular level in the style hierarchy (typically via direct formatting).
            if (numId == 0)
                return NumberingFormat.Removed;

            // Find the number definition instance.
            var styles = p.Document.NumberingStyles;
            var definition = styles.Definitions.SingleOrDefault(d => d.Id == numId);
            if (definition == null)
            {
                throw new Exception(
                    $"Number reference w:numId('{numId}') used in document but not defined in /word/numbering.xml");
            }

            // Get the level
            int level = int.Parse(numProperties.Element(Namespace.Main + "ilvl").GetVal());
            return definition.Style.Levels
                .Single(l => l.Level == level)
                .Format;
        }
    }
}