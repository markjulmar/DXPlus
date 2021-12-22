using DXPlus.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a single run of text with optional formatting in a paragraph
    /// </summary>
    [DebuggerDisplay("{" + nameof(Text) + "}")]
    public class Run : IEquatable<Run>
    {
        /// <summary>
        /// Document for this run - used to retrieve images.
        /// </summary>
        internal IDocument Document { get; }

        /// <summary>
        /// XML backing storage
        /// </summary>
        internal XElement Xml { get; }

        /// <summary>
        /// Gets the start index of this Text (text length before this text)
        /// </summary>
        public int StartIndex { get; }

        /// <summary>
        /// Gets the end index of this Text (text length before this text + this texts length)
        /// </summary>
        public int EndIndex { get; }

        /// <summary>
        /// True if this run has a text block. False if it's a linebreak, paragraph break, or empty.
        /// </summary>
        public bool HasText => Xml.Element(Name.Text) != null;
        
        /// <summary>
        /// The formatted text value of this run
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Returns the breaks in this run
        /// </summary>
        public IEnumerable<TextElement> Elements
            => Xml.Elements()
                .Where(e => e.Name != Name.RunProperties)
                .Select(e => WrapTextChild(e));

        /// <summary>
        /// Wraps a child element in an accessor object.
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        private TextElement WrapTextChild(XElement child)
        {
            return child.Name.LocalName switch
            {
                "br" => new Break(this, child),
                "t" => new Text(this, child),
                "drawing" => new Drawing(this, child),
                "commentReference" => new CommentRef(this, child),
                _ => new TextElement(this, child),
            };
        }

        /// <summary>
        /// Style applied to this run
        /// </summary>
        public string StyleName
        {
            get => Xml.GetRunProps(false)?.Element(Namespace.Main + "rStyle")?.GetVal();
            set => Xml.GetRunProps(true).AddElementVal(Namespace.Main + "rStyle", value);
        }

        /// <summary>
        /// The run properties for this text run
        /// </summary>
        public Formatting Properties
        {
            get => new(Xml.GetRunProps(true));
            set
            {
                XElement rPr = Xml.Element(Name.RunProperties);
                rPr?.Remove();

                if (value != null)
                {
                    var xml = value.Xml;
                    if (xml.Parent != null)
                        xml = xml.Clone();
                    Xml.AddFirst(xml);
                }
            }
        }

        /// <summary>
        /// Add/Remove the specific formatting specified from this run.
        /// </summary>
        /// <param name="other">Formatting to apply</param>
        public void AddFormatting(Formatting other) => Properties.Merge(other);

        /// <summary>
        /// Constructor for a run of text
        /// </summary>
        /// <param name="document">Document</param>
        /// <param name="xml"></param>
        /// <param name="startIndex"></param>
        internal Run(IDocument document, XElement xml, int startIndex)
        {
            Document = document; // can be null.
            Xml = xml ?? throw new ArgumentNullException(nameof(xml));
            StartIndex = startIndex;
            int currentPos = startIndex;

            // Determine the end and get the raw text from the run.
            foreach (var te in xml.Descendants())
            {
                var text = HelperFunctions.ToText(te);
                if (!string.IsNullOrEmpty(text))
                {
                    Text += text;
                    currentPos += text.Length;
                }
            }
            EndIndex = currentPos;
        }

        /// <summary>
        /// Split a run at a given index.
        /// </summary>
        /// <param name="index">Index to split this run at</param>
        /// <param name="editType">Type of editing being performed</param>
        /// <returns></returns>
        internal XElement[] SplitAtIndex(int index, EditType editType = EditType.Insert)
        {
            // Find the (w:t) we need to split based on the index.
            index -= StartIndex;
            var (textXml, startIndex) = FindTextElementByIndex(Xml, index);

            // Split the block.
            // Returns [textElement, leftSide, rightSide]
            var splitText = SplitTextElementAtIndex(index, textXml, startIndex);
            
            var splitLeft = new XElement(Xml.Name,
                                        Xml.Attributes(),
                                        Xml.Element(Name.RunProperties),
                                        splitText[0].ElementsBeforeSelf().Where(n => n.Name != Name.RunProperties),
                                        splitText[1]);

            if (HelperFunctions.GetTextLength(splitLeft) == 0)
            {
                splitLeft = null;
            }

            var splitRight = new XElement(Xml.Name,
                                        Xml.Attributes(),
                                        Xml.Element(Name.RunProperties),
                                        splitText[2],
                                        splitText[0].ElementsAfterSelf().Where(n => n.Name != Name.RunProperties));

            if (HelperFunctions.GetTextLength(splitRight) == 0)
            {
                splitRight = null;
            }

            return new[] { splitLeft, splitRight };
        }

        /// <summary>
        /// Split the text block at the given index
        /// </summary>
        /// <param name="index">Index to split at in parent Run</param>
        /// <param name="xml">Text block to split</param>
        /// <param name="startIndex">Start index of the text block in parent Run</param>
        /// <returns>Array with left/right XElement values</returns>
        private static XElement[] SplitTextElementAtIndex(int index, XElement xml, int startIndex)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));

            int endIndex = startIndex + HelperFunctions.GetSize(xml);
            if (index < startIndex || index > endIndex)
                throw new ArgumentOutOfRangeException(nameof(index));

            XElement splitLeft = null, splitRight = null;

            if (xml.Name.LocalName == "t" || xml.Name.LocalName == "delText")
            {
                // The original text element, now containing only the text before the index point.
                splitLeft = new XElement(xml.Name, xml.Attributes(), xml.Value.Substring(0, index - startIndex));
                if (splitLeft.Value.Length == 0)
                {
                    splitLeft = null;
                }
                else
                {
                    splitLeft.PreserveSpace();
                }

                // The original text element, now containing only the text after the index point.
                splitRight = new XElement(xml.Name, xml.Attributes(), xml.Value.Substring(index - startIndex, xml.Value.Length - (index-startIndex)));
                if (splitRight.Value.Length == 0)
                {
                    splitRight = null;
                }
                else
                {
                    splitRight.PreserveSpace();
                }
            }
            else
            {
                if (index == endIndex)
                {
                    splitLeft = xml;
                }
                else
                {
                    splitRight = xml;
                }
            }

            return new[] { xml, splitLeft, splitRight };
        }

        /// <summary>
        /// Internal method to recursively walk all (w:t) elements in this run and find the spot
        /// where an edit (insert/delete) would occur.
        /// </summary>
        /// <param name="element">XML graph to examine</param>
        /// <param name="index">Index to search for</param>
        private static (XElement textXml, int startIndex) FindTextElementByIndex(XElement element, int index)
        {
            int count = 0;
            foreach (var child in element.Descendants())
            {
                int size = HelperFunctions.GetSize(child);
                count += size;
                if (count >= index)
                {
                    return (child, count - size);
                }
            }
            return default;
        }

        /// <summary>
        /// Determines equality for a run
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Run other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
    }
}