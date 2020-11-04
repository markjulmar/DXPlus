using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Style definition which groups all the style properties.
    /// </summary>
    public sealed class Style
    {
        internal XElement Xml { get; }

        /// <summary>
        /// Unique id for this style
        /// </summary>
        public string Id
        {
            get => Xml.AttributeValue(Namespace.Main + "styleId");
            set => Xml.SetAttributeValue(Namespace.Main + "styleId", value);
        }

        /// <summary>
        /// Name of the style
        /// </summary>
        public string Name
        {
            get => Xml.Element(DXPlus.Name.NameId).GetVal();
            set => Xml.AddElementVal(DXPlus.Name.NameId, value);
        }

        /// <summary>
        /// Style is a user-defined style.
        /// </summary>
        public bool IsCustom
        {
            get => Xml.BoolAttributeValue(Namespace.Main + "customStyle");
            set => Xml.SetAttributeValue(Namespace.Main + "customStyle", value ? "1" : null);
        }

        /// <summary>
        /// Specifies that this style is the default for the given Type.
        /// </summary>
        public bool IsDefault
        {
            get => Xml.BoolAttributeValue(Namespace.Main + "default");
            set => Xml.SetAttributeValue(Namespace.Main + "default", value ? "1" : null);
        }

        /// <summary>
        /// The type this style is applied to
        /// </summary>
        public StyleType Type
        {
            get => Xml.AttributeValue(Namespace.Main + "type").TryGetEnumValue<StyleType>(out var result)
                ? result
                : StyleType.Paragraph;

            set => Xml.SetAttributeValue(Namespace.Main + "type", value.GetEnumName());
        }

        /// <summary>
        /// Retrieve the formatting options
        /// </summary>
        public Formatting Formatting => new Formatting(Xml.GetOrCreateElement(DXPlus.Name.RunProperties));

        /// <summary>
        /// Paragraph properties
        /// </summary>
        public ParagraphProperties ParagraphFormatting => new ParagraphProperties(Xml.GetOrCreateElement(DXPlus.Name.ParagraphProperties));

        /// <summary>
        /// The style this one is based on.
        /// </summary>
        public string BasedOn
        {
            get => Xml.Element(Namespace.Main + "basedOn").GetVal(null);
            set => Xml.AddElementVal(Namespace.Main + "basedOn", string.IsNullOrWhiteSpace(value) ? null : value);
        }

        /// <summary>
        /// The default style for the next paragraph
        /// </summary>
        public string NextParagraphStyle
        {
            get => Xml.Element(Namespace.Main + "next").GetVal(null);
            set => Xml.AddElementVal(Namespace.Main + "next", string.IsNullOrWhiteSpace(value) ? null : value);
        }

        /// <summary>
        /// Linked style
        /// </summary>
        public string Linked
        {
            get => Xml.Element(Namespace.Main + "link").GetVal(null);
            set => Xml.AddElementVal(Namespace.Main + "link", string.IsNullOrWhiteSpace(value) ? null : value);
        }

        // TODO: add tblPr, tblStylePr, tcPr, trPr

        /// <summary>
        /// Constructor for an existing style
        /// </summary>
        /// <param name="xml">Element in the style document</param>
        public Style(XElement xml)
        {
            Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        }
    }
}