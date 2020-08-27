using System;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a single run of text in a paragraph
    /// </summary>
    public class Run : DocXBase
    {
        /// <summary>
        /// Gets the start index of this Text (text length before this text)
        /// </summary>
        public int StartIndex { get; }

        /// <summary>
        /// Gets the end index of this Text (text length before this text + this texts length)
        /// </summary>
        public int EndIndex { get; }

        /// <summary>
        /// The text value of this text element
        /// </summary>
        internal string Value { set; get; }

        /// <summary>
        /// Constructor for a run of text
        /// </summary>
        /// <param name="document"></param>
        /// <param name="xml"></param>
        /// <param name="startIndex"></param>
        internal Run(IDocument document, XElement xml, int startIndex)
            : base(document, xml)
        {
            StartIndex = startIndex;
            int currentPos = startIndex;

            // Loop through each text in this run
            foreach (var te in xml.Descendants())
            {
                switch (te.Name.LocalName)
                {
                    case "tab":
                        Value += "\t";
                        currentPos++;
                        break;

                    case "br":
                        Value += "\n";
                        currentPos++;
                        break;

                    case "t":
                    case "delText":
                        if (te.Value.Length > 0)
                        {
                            Value += te.Value;
                            currentPos += te.Value.Length;
                        }
                        break;
                }
            }

            EndIndex = currentPos;
        }

        internal static XElement[] SplitRun(Run run, int index, EditType editType = EditType.Ins)
        {
            index -= run.StartIndex;

            TextBlock text = run.GetFirstTextEffectedByEdit(index, editType);
            
            var splitText = TextBlock.SplitText(text, index);

            var splitLeft = new XElement(run.Xml.Name,
                                        run.Xml.Attributes(),
                                        run.Xml.GetRunProps(false),
                                        text.Xml.ElementsBeforeSelf().Where(n => n.Name.LocalName != "rPr"),
                                        splitText[0]);

            if (Paragraph.GetElementTextLength(splitLeft) == 0)
            {
                splitLeft = null;
            }

            var splitRight = new XElement(run.Xml.Name,
                                        run.Xml.Attributes(),
                                        run.Xml.GetRunProps(false),
                                        splitText[1],
                                        text.Xml.ElementsAfterSelf().Where(n => n.Name.LocalName != "rPr"));

            if (Paragraph.GetElementTextLength(splitRight) == 0)
            {
                splitRight = null;
            }

            return new[] { splitLeft, splitRight };
        }

        internal TextBlock GetFirstTextEffectedByEdit(int index, EditType type = EditType.Ins)
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

        internal void GetFirstTextEffectedByEditRecursive(XElement element, int index, ref int count, ref TextBlock theOne, EditType type = EditType.Ins)
        {
            count += HelperFunctions.GetSize(element);
            if (count > 0 && ((type == EditType.Del && count > index) || type == EditType.Ins && count >= index))
            {
                theOne = new TextBlock(Document, element, count - HelperFunctions.GetSize(element));
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