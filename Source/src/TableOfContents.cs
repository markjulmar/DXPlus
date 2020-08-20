using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a table of contents in the document
    /// </summary>
    public sealed class TableOfContents : DocXElement
    {
        private const string HeaderStyle = "TOCHeading";
        private const int RightTabPos = 9350;

        private TableOfContents(DocX document, XElement xml, string headerStyle) : base(document, xml)
        {
            if (!document.settingsDoc.Descendants(DocxNamespace.Main + "updateFields").Any())
            {
                // Tell Word to update the document ToC the next time this document is loaded.
                document.settingsDoc.Root.Add(new XElement(DocxNamespace.Main + "updateFields",
                                           new XAttribute(DocxNamespace.Main + "val", true)));
            }

            EnsureTocStylesArePresent(document, headerStyle);
        }

        internal static TableOfContents CreateTableOfContents(DocX document, string title, TableOfContentsSwitches switches, string headerStyle = null, int lastIncludeLevel = 3, int? rightTabPos = null)
        {
            XElement tocBase = Resources.TocXmlBase(
                                    headerStyle ?? HeaderStyle,
                                    title, rightTabPos ?? RightTabPos,
                                    BuildSwitchString(switches, lastIncludeLevel));
            return new TableOfContents(document, tocBase, headerStyle);
        }

        private static string BuildSwitchString(TableOfContentsSwitches switches, int lastIncludeLevel)
        {
            string switchString = "TOC";

            System.Collections.Generic.IEnumerable<TableOfContentsSwitches> allSwitches = Enum.GetValues(typeof(TableOfContentsSwitches)).Cast<TableOfContentsSwitches>();
            foreach (TableOfContentsSwitches tocSwitch in allSwitches.Where(s => s != TableOfContentsSwitches.None && (switches & s) != 0))
            {
                switchString += " " + tocSwitch.GetEnumName();
                if (tocSwitch == TableOfContentsSwitches.O)
                {
                    switchString += $" '{1}-{lastIncludeLevel}'";
                }
            }

            return switchString;
        }

        private void EnsureTocStylesArePresent(DocX document, string headerStyle)
        {
            (string headerStyle, string applyTo, Func<string, string, XElement> template, string name)[] availableStyles = new (string headerStyle, string applyTo, Func<string, string, XElement> template, string name)[]
            {
                (headerStyle, "paragraph", Resources.TocHeadingStyleBase, headerStyle ?? HeaderStyle),
                ("TOC1", "paragraph", Resources.TocElementStyleBase, "toc 1"),
                ("TOC2", "paragraph", Resources.TocElementStyleBase, "toc 2"),
                ("TOC3", "paragraph", Resources.TocElementStyleBase, "toc 3"),
                ("TOC4", "paragraph", Resources.TocElementStyleBase, "toc 4"),
                ("Hyperlink", "character", Resources.TocHyperLinkStyleBase, "")
            };

            foreach (var style in availableStyles)
            {
                if (!document.HasStyle(style.headerStyle, style.applyTo))
                {
                    XElement xml = style.template.Invoke(style.headerStyle, style.name);
                    document.stylesDoc.Root.Add(xml);
                }
            }
        }
    }
}