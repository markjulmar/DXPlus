using System;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    internal class TextBlock : DocXBase
    {
        internal TextBlock(IDocument document, XElement xml, int startIndex) : base(document, xml)
        {
            StartIndex = startIndex;

            switch (Xml.Name.LocalName)
            {
                case "t":
                case "delText":
                    EndIndex = startIndex + xml.Value.Length;
                    Value = xml.Value;
                    break;

                case "br":
                    Value = "\n";
                    EndIndex = startIndex + 1;
                    break;

                case "tab":
                    Value = "\t";
                    EndIndex = startIndex + 1;
                    break;
            }
        }

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
        public string Value { get; }

        internal static XElement[] SplitText(TextBlock t, int index)
        {
            if (index < t.StartIndex || index > t.EndIndex)
                throw new ArgumentOutOfRangeException(nameof(index));

            XElement splitLeft = null, splitRight = null;

            if (t.Xml.Name.LocalName == "t" || t.Xml.Name.LocalName == "delText")
            {
                // The original text element, now containing only the text before the index point.
                splitLeft = new XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(0, index - t.StartIndex));
                if (splitLeft.Value.Length == 0)
                {
                    splitLeft = null;
                }
                else
                {
                    splitLeft.PreserveSpace();
                }

                // The original text element, now containing only the text after the index point.
                splitRight = new XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(index - t.StartIndex, t.Xml.Value.Length - (index - t.StartIndex)));
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
                if (index == t.EndIndex)
                {
                    splitLeft = t.Xml;
                }
                else
                {
                    splitRight = t.Xml;
                }
            }

            return new[] { splitLeft, splitRight };
        }
    }
}