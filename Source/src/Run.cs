using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    public class Run : DocXElement
    {
        // A lookup for the text elements in this paragraph
        private readonly Dictionary<int, TextBlock> textLookup = new Dictionary<int, TextBlock>();

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

        internal Run(DocX document, XElement xml, int startIndex)
            : base(document, xml)
        {
            StartIndex = startIndex;
            int currentPos = startIndex;

            // Loop through each text in this run
            foreach (XElement te in xml.Descendants())
            {
                switch (te.Name.LocalName)
                {
                    case "tab":
                        textLookup.Add(currentPos + 1, new TextBlock(Document, te, currentPos));
                        Value += "\t";
                        currentPos++;
                        break;

                    case "br":
                        textLookup.Add(currentPos + 1, new TextBlock(Document, te, currentPos));
                        Value += "\n";
                        currentPos++;
                        break;

                    case "t": goto case "delText";
                    case "delText":
                        // Only add strings which are not empty
                        if (te.Value.Length > 0)
                        {
                            textLookup.Add(currentPos + te.Value.Length, new TextBlock(Document, te, currentPos));
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
            XElement[] splitText = TextBlock.SplitText(text, index);

            XElement splitLeft = new XElement(run.Xml.Name,
                                        run.Xml.Attributes(),
                                        run.Xml.GetRunProps(false),
                                        text.Xml.ElementsBeforeSelf().Where(n => n.Name.LocalName != "rPr"),
                                        splitText[0]);

            if (Paragraph.GetElementTextLength(splitLeft) == 0)
            {
                splitLeft = null;
            }

            XElement splitRight = new XElement(run.Xml.Name,
                                        run.Xml.Attributes(),
                                        run.Xml.GetRunProps(false),
                                        splitText[1],
                                        text.Xml.ElementsAfterSelf().Where(n => n.Name.LocalName != "rPr"));

            if (Paragraph.GetElementTextLength(splitRight) == 0)
            {
                splitRight = null;
            }

            return new XElement[] { splitLeft, splitRight };
        }

        internal TextBlock GetFirstTextEffectedByEdit(int index, EditType type = EditType.Ins)
        {
            // Make sure we are looking within an acceptable index range.
            if (index < 0 || index > HelperFunctions.GetText(Xml).Length)
            {
                throw new ArgumentOutOfRangeException();
            }

            // Need some memory that can be updated by the recursive search for the XElement to Split.
            int count = 0;
            TextBlock theOne = null;

            GetFirstTextEffectedByEditRecursive(Xml, index, ref count, ref theOne, type);

            return theOne;
        }

        internal void GetFirstTextEffectedByEditRecursive(XElement Xml, int index, ref int count, ref TextBlock theOne, EditType type = EditType.Ins)
        {
            count += HelperFunctions.GetSize(Xml);
            if (count > 0 && ((type == EditType.Del && count > index) || (type == EditType.Ins && count >= index)))
            {
                theOne = new TextBlock(Document, Xml, count - HelperFunctions.GetSize(Xml));
                return;
            }

            if (Xml.HasElements)
            {
                foreach (XElement e in Xml.Elements())
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