using System;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a single run of text in a paragraph
    /// </summary>
    [DebuggerDisplay("{Text}")]
    internal class Run
    {
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
        /// The raw text value of this text element
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// The run properties for this text run
        /// </summary>
        public Formatting Formatting => new Formatting(Xml.Element(Name.RunProperties));

        /// <summary>
        /// Constructor for a run of text
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="startIndex"></param>
        internal Run(XElement xml, int startIndex)
        {
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

        internal XElement[] SplitRun(int index, EditType editType = EditType.Insert)
        {
            index -= StartIndex;

            TextBlock text = GetFirstTextAffectedByEdit(index, editType);
            
            var splitText = text.Split(index);

            var splitLeft = new XElement(Xml.Name,
                                        Xml.Attributes(),
                                        Xml.Element(Name.RunProperties),
                                        text.Xml.ElementsBeforeSelf().Where(n => n.Name != Name.RunProperties),
                                        splitText[0]);

            if (Paragraph.GetElementTextLength(splitLeft) == 0)
            {
                splitLeft = null;
            }

            var splitRight = new XElement(Xml.Name,
                                        Xml.Attributes(),
                                        Xml.Element(Name.RunProperties),
                                        splitText[1],
                                        text.Xml.ElementsAfterSelf().Where(n => n.Name != Name.RunProperties));

            if (Paragraph.GetElementTextLength(splitRight) == 0)
            {
                splitRight = null;
            }

            return new[] { splitLeft, splitRight };
        }

        internal TextBlock GetFirstTextAffectedByEdit(int index, EditType type = EditType.Insert)
        {
            // Make sure we are looking within an acceptable index range.
            if (index < 0 || index > HelperFunctions.GetText(Xml).Length)
                throw new ArgumentOutOfRangeException(nameof(index));

            // Need some memory that can be updated by the recursive search for the XElement to Split.
            int count = 0;
            TextBlock theOne = null;

            GetFirstTextEffectedByEditRecursive(Xml, index, ref count, ref theOne, type);

            return theOne;
        }

        internal void GetFirstTextEffectedByEditRecursive(XElement element, int index, ref int count, ref TextBlock theOne, EditType type = EditType.Insert)
        {
            count += HelperFunctions.GetSize(element);
            if (count > 0 && ((type == EditType.Delete && count > index) || (type == EditType.Insert && count >= index)))
            {
                theOne = new TextBlock(element, count - HelperFunctions.GetSize(element));
                return;
            }

            if (element.HasElements)
            {
                foreach (var e in element.Elements())
                {
                    if (theOne == null)
                    {
                        GetFirstTextEffectedByEditRecursive(e, index, ref count, ref theOne);
                    }
                }
            }
        }
    }
}