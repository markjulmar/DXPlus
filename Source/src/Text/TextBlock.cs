using DXPlus.Helpers;
using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This represents a piece of text (typically a 'w:t') in a Run (w:r).
    /// </summary>
    internal class TextBlock
    {
        /// <summary>
        /// XML fragment in parent Run
        /// </summary>
        internal XElement Xml { get; }

        /// <summary>
        /// The text value of this text element
        /// </summary>
        public string Value { get; }

        public int StartIndex { get; }
        public int EndIndex { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="xml">XML fragment with a single element from parent run</param>
        /// <param name="startIndex">Starting index in parent run</param>
        internal TextBlock(XElement xml, int startIndex)
        {
            StartIndex = startIndex;
            Xml = xml ?? throw new ArgumentNullException(nameof(xml));

            switch (Xml.Name.LocalName)
            {
                case "t":
                case "delText":
                    Value = xml.Value;
                    break;

                case "cr":
                case "br":
                    Value = "\n";
                    break;

                case "tab":
                    Value = "\t";
                    break;

                default:
                    Value = string.Empty;
                    break;
            }

            EndIndex = startIndex + Value.Length;
        }

        /// <summary>
        /// Split the text block at the given index
        /// </summary>
        /// <param name="index">Index to split at in parent Run values</param>
        /// <returns>Array with left/right XElement values</returns>
        internal XElement[] Split(int index)
        {
            if (index < StartIndex || index > EndIndex)
                throw new ArgumentOutOfRangeException(nameof(index));

            XElement splitLeft = null, splitRight = null;

            if (Xml.Name.LocalName == "t" || Xml.Name.LocalName == "delText")
            {
                // The original text element, now containing only the text before the index point.
                splitLeft = new XElement(Xml.Name, Xml.Attributes(), Xml.Value.Substring(0, index - StartIndex));
                if (splitLeft.Value.Length == 0)
                {
                    splitLeft = null;
                }
                else
                {
                    splitLeft.PreserveSpace();
                }

                // The original text element, now containing only the text after the index point.
                splitRight = new XElement(Xml.Name, Xml.Attributes(), Xml.Value.Substring(index - StartIndex, Xml.Value.Length - (index-StartIndex)));
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
                if (index == EndIndex)
                {
                    splitLeft = Xml;
                }
                else
                {
                    splitRight = Xml;
                }
            }

            return new[] { splitLeft, splitRight };
        }
    }
}