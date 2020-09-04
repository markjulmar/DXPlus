using DXPlus.Helpers;
using System;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a single run of text with optional formatting in a paragraph
    /// </summary>
    [DebuggerDisplay("{" + nameof(Text) + "}")]
    public class Run
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
        /// The raw text value of this run
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// The run properties for this text run
        /// </summary>
        public Formatting Properties => new Formatting(Xml.GetRunProps(true));

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

        /// <summary>
        /// Split a run at a given index.
        /// </summary>
        /// <param name="index">Index to split this run at</param>
        /// <param name="editType">Type of editing being performed</param>
        /// <returns></returns>
        internal XElement[] SplitRun(int index, EditType editType = EditType.Insert)
        {
            index -= StartIndex;

            TextBlock text = FindTextAffectedByEdit(editType, index);
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

        /// <summary>
        /// Walk the run and identify the first (w:t) text element affected by an insert or delete
        /// </summary>
        /// <param name="type">Type of edit being performed</param>
        /// <param name="index">Position of edit</param>
        /// <returns></returns>
        private TextBlock FindTextAffectedByEdit(EditType type, int index)
        {
            // Make sure we are looking within an acceptable index range.
            if (index < 0 || index > HelperFunctions.GetText(Xml).Length)
                throw new ArgumentOutOfRangeException(nameof(index));

            // Start the recursive search
            int count = 0; TextBlock locatedBlock = null;
            RecursiveSearchForTextByIndex(Xml, type, index, ref count, ref locatedBlock);

            return locatedBlock;
        }

        /// <summary>
        /// Internal method to recursively walk all (w:t) elements in this run and find the spot
        /// where an edit (insert/delete) would occur.
        /// </summary>
        /// <param name="element">XML graph to examine</param>
        /// <param name="type">Insert or delete</param>
        /// <param name="index">Position where edit is being performed</param>
        /// <param name="count">Number of text characters encountered so far</param>
        /// <param name="textBlock">The identified text block</param>
        private static void RecursiveSearchForTextByIndex(XElement element, EditType type, int index, ref int count, ref TextBlock textBlock)
        {
            count += HelperFunctions.GetSize(element);
            if (count > 0
                && ((type == EditType.Delete && count > index)
                    || (type == EditType.Insert && count >= index)))
            {
                textBlock = new TextBlock(element, count - HelperFunctions.GetSize(element));
                return;
            }

            if (element.HasElements)
            {
                foreach (var e in element.Elements())
                {
                    RecursiveSearchForTextByIndex(e, EditType.Insert, index, ref count, ref textBlock);
                    if (textBlock != null)
                        return;
                }
            }
        }
    }
}