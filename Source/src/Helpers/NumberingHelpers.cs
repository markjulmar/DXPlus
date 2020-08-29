using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    internal static class NumberingHelpers
    {
        /// <summary>
        /// Determine if this paragraph is a list element.
        /// </summary>
        internal static bool IsListItem(this Paragraph p) => ParagraphNumberProperties(p) != null;

        /// <summary>
        /// Fetch the paragraph number properties for a list element.
        /// </summary>
        internal static XElement ParagraphNumberProperties(this Paragraph p) => p.Xml.FirstLocalNameDescendant("numPr");

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
                : int.Parse(numProperties.Element(Namespace.Main + "numId").GetVal());
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
                : int.Parse(numProperties.Element(Namespace.Main + "ilvl").GetVal());
        }

        /// <summary>
        /// Get the ListItemType property for the paragraph.
        /// Defaults to numbered if a list is found but the type is not specified
        /// </summary>
        internal static NumberingFormat GetNumberingFormat(this Paragraph p)
        {
            var numProperties = ParagraphNumberProperties(p);
            if (numProperties == null)
                return NumberingFormat.None;

            int level = int.Parse(numProperties.Element(Namespace.Main + "ilvl").GetVal());
            int numId = int.Parse(numProperties.Element(Namespace.Main + "numId").GetVal());

            // Find the number definition instance.
            var styles = p.Document.NumberingStyles;
            var definition = styles.Definitions.SingleOrDefault(d => d.Id == numId);
            if (definition == null)
            {
                throw new Exception(
                    $"Number reference w:numId('{numId}') used in document but not defined in /word/numbering.xml");
            }

            return definition.Style.Levels
                .Single(l => l.Level == level)
                .Format;
        }
    }
}
