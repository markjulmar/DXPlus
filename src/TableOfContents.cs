using DXPlus.Resources;
using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a table of contents in the document
    /// </summary>
    public sealed class TableOfContents : DocXElement
    {
        private const string HeaderStyle = "TOCHeading";
        private const int RightTabPos = 9350;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML data</param>
        private TableOfContents(IDocument document, XElement xml) : base(document, xml)
        {
        }

        /// <summary>
        /// Create a TOC
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="title">Title</param>
        /// <param name="switches">TOC switches</param>
        /// <param name="headerStyle">Header style (null for default style)</param>
        /// <param name="lastIncludeLevel">Last level to include</param>
        /// <param name="rightTabPos">Position of right tab</param>
        internal static TableOfContents CreateTableOfContents(Document document, string title,
                                                    TableOfContentsSwitches switches, string headerStyle = null,
                                                    int lastIncludeLevel = 3, int? rightTabPos = null)
        {
            headerStyle ??= HeaderStyle;
            rightTabPos ??= RightTabPos;
            title ??= string.Empty;

            // Invalidate placeholder fields
            document.InvalidatePlaceholderFields();

            // Add any required styles to the document
            EnsureTocStylesArePresent(document, headerStyle);

            // Create the TOC
            return new TableOfContents(document,
                Resource.TocXmlBase(headerStyle, title, rightTabPos,
                    BuildSwitchString(switches, lastIncludeLevel)));
        }

        /// <summary>
        /// Build the text TOC switches set on this table of contents.
        /// </summary>
        private static string BuildSwitchString(TableOfContentsSwitches switches, int lastIncludeLevel)
        {
            var switchString = "TOC";

            var allSwitches = Enum.GetValues(typeof(TableOfContentsSwitches)).Cast<TableOfContentsSwitches>();
            foreach (var tocSwitch in allSwitches.Where(s => s != TableOfContentsSwitches.None && (switches & s) != 0))
            {
                switchString += " " + tocSwitch.GetEnumName();
                if (tocSwitch == TableOfContentsSwitches.O)
                {
                    switchString += $" '{1}-{lastIncludeLevel}'";
                }
            }

            return switchString;
        }

        /// <summary>
        /// Add any required missing TOC styles to the passed document.
        /// </summary>
        /// <param name="document">Document to alter</param>
        /// <param name="headerStyle">Header style, null to use the default style</param>
        private static void EnsureTocStylesArePresent(IDocument document, string headerStyle)
        {
            headerStyle ??= HeaderStyle;

            var availableStyles = new (string headerStyle, StyleType applyTo, Func<string, string, XElement> template, string name)[]
            {
                (headerStyle, StyleType.Paragraph, Resource.TocHeadingStyleBase, headerStyle),
                ("TOC1", StyleType.Paragraph, Resource.TocElementStyleBase, "toc 1"),
                ("TOC2", StyleType.Paragraph, Resource.TocElementStyleBase, "toc 2"),
                ("TOC3", StyleType.Paragraph, Resource.TocElementStyleBase, "toc 3"),
                ("TOC4", StyleType.Paragraph, Resource.TocElementStyleBase, "toc 4"),
                ("Hyperlink", StyleType.Character, Resource.TocHyperLinkStyleBase, "")
            };

            var mgr = document.Styles;
            foreach (var (style, applyTo, template, name) in availableStyles)
            {
                if (!mgr.HasStyle(style, applyTo))
                {
                    var xml = template.Invoke(style, name);
                    mgr.Add(xml);
                }
            }
        }
    }
}