using System;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a single level in a numbering definition style.
    /// </summary>
    public sealed class NumberingLevel
    {
        /// <summary>
        /// The XML fragment making up this object (w:lvl)
        /// </summary>
        public XElement Xml { get; }

        /// <summary>
        /// Level this definition is associated with
        /// </summary>
        public int Level
        {
            get => int.Parse(Xml.AttributeValue(Namespace.Main + "ilvl"));
            set => Xml.SetAttributeValue(Namespace.Main + "ilvl", value);
        }

        /// <summary>
        /// The starting value for this level.
        /// </summary>
        public int Start
        {
            get => int.Parse(Xml.Element(Namespace.Main + "start")?.GetVal() ?? "0");
            set => Xml.AddElementVal(Namespace.Main + "start", value);
        }

        /// <summary>
        /// Retrieve the formatting options
        /// </summary>
        public Formatting Formatting => new Formatting(Xml.GetOrCreateElement(Name.RunProperties));

        /// <summary>
        /// Paragraph properties
        /// </summary>
        public ParagraphProperties ParagraphFormatting => new ParagraphProperties(Xml.GetOrCreateElement(Name.ParagraphProperties));

        /// <summary>
        /// Number format used to display all the values at this level.
        /// </summary>
        public NumberingFormat Format
        {
            get => Xml.Element(Namespace.Main + "numFmt")
                .GetVal().TryGetEnumValue<NumberingFormat>(out var result) ? result : NumberingFormat.None;
            set => Xml.AddElementVal(Namespace.Main + "numFmt", value.GetEnumName());
        }

        /// <summary>
        /// Specifies the content to be displayed when displaying a paragraph at this level.
        /// </summary>
        public string Text
        {
            get => Xml.Element(Namespace.Main + "lvlText").GetVal();
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Value must be provided.", nameof(Text));
                Xml.GetOrCreateElement(Namespace.Main + "lvlText").SetAttributeValue(Name.MainVal, value);
            }
        }

        public Alignment Alignment
        {
            get => Xml.Element(Namespace.Main + "lvlJc")
                .GetVal().TryGetEnumValue<Alignment>(out var result) ? result : Alignment.Left;
            set => Xml.AddElementVal(Namespace.Main + "lvlJc", value.GetEnumName());
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="xml"></param>
        internal NumberingLevel(XElement xml)
        {
            this.Xml = xml;
        }
    }
}