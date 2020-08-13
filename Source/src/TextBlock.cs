using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    internal class TextBlock : DocXElement
    {
        internal TextBlock(DocX document, XElement xml, int startIndex)
                    : base(document, xml)
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

        /// <summary>
        /// If a text element or delText element, starts or ends with a space,
        /// it must have the attribute space, otherwise it must not have it.
        /// </summary>
        /// <param name="e">The (t or delText) element check</param>
        public static void PreserveSpace(XElement e)
        {
            if (!e.Name.Equals(DocxNamespace.Main + "t")
             && !e.Name.Equals(DocxNamespace.Main + "delText"))
            {
                throw new ArgumentException($"{nameof(PreserveSpace)} can only work with elements of type 't' or 'delText'", nameof(e));
            }

            // Check if this w:t contains a space atribute
            XAttribute space = e.Attributes().SingleOrDefault(a => a.Name.Equals(XNamespace.Xml + "space"));

            // This w:t's text begins or ends with whitespace
            if (e.Value.StartsWith(" ") || e.Value.EndsWith(" "))
            {
                // If this w:t contains no space attribute, add one.
                if (space == null)
                    e.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
            }

            // This w:t's text does not begin or end with a space
            else
            {
                // If this w:r contains a space attribute, remove it.
                space?.Remove();
            }
        }

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
                    splitLeft = null;
                else
                    PreserveSpace(splitLeft);

                // The original text element, now containing only the text after the index point.
                splitRight = new XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(index - t.StartIndex, t.Xml.Value.Length - (index - t.StartIndex)));
                if (splitRight.Value.Length == 0)
                    splitRight = null;
                else
                    PreserveSpace(splitRight);
            }
            else
            {
                if (index == t.EndIndex)
                    splitLeft = t.Xml;
                else
                    splitRight = t.Xml;
            }

            return new XElement[] { splitLeft, splitRight };
        }
    }
}