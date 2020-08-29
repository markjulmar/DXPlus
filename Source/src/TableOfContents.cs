using DXPlus.Helpers;
using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a table of contents in the document
    /// </summary>
    public sealed class TableOfContents : DocXBase
    {
        private const string HeaderStyle = "TOCHeading";
        private const int RightTabPos = 9350;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML data</param>
        /// <param name="headerStyle">Header style</param>
        private TableOfContents(DocX document, XElement xml, string headerStyle) : base(document, xml)
        {
            // Invalidate placeholder fields
            document.InvalidatePlaceholderFields();
            
            // Add any required styles to the document
            EnsureTocStylesArePresent(document, headerStyle);
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
        internal static TableOfContents CreateTableOfContents(DocX document, string title,
                                                    TableOfContentsSwitches switches, string headerStyle = null,
                                                    int lastIncludeLevel = 3, int? rightTabPos = null)
            => new TableOfContents(document,
                Resources.TocXmlBase(
                    headerStyle ?? HeaderStyle,
                    title,
                    rightTabPos ?? RightTabPos,
                    BuildSwitchString(switches, lastIncludeLevel)),
                headerStyle);

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
        private static void EnsureTocStylesArePresent(DocX document, string headerStyle)
        {
            headerStyle ??= HeaderStyle;

            var availableStyles = new (string headerStyle, StyleType applyTo, Func<string, string, XElement> template, string name)[]
            {
                (headerStyle, StyleType.Paragraph, Resources.TocHeadingStyleBase, headerStyle),
                ("TOC1", StyleType.Paragraph, Resources.TocElementStyleBase, "toc 1"),
                ("TOC2", StyleType.Paragraph, Resources.TocElementStyleBase, "toc 2"),
                ("TOC3", StyleType.Paragraph, Resources.TocElementStyleBase, "toc 3"),
                ("TOC4", StyleType.Paragraph, Resources.TocElementStyleBase, "toc 4"),
                ("Hyperlink", StyleType.Character, Resources.TocHyperLinkStyleBase, "")
            };

            document.AddDefaultStyles();
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