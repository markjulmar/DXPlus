using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Creates an underline style for a Text Run.
    /// </summary>
    public sealed class Underline
    {
        private readonly XElement? properties;
        private XElement? xml;

        internal XElement Xml
        {
            get
            {
                if (xml == null)
                {
                    xml = new XElement(Name.Underline);
                    properties?.Element(Name.Underline)?.Remove();
                    properties?.Add(xml);
                }
                return xml;
            }
        }

        /// <summary>
        /// Converter from bool to Underline
        /// </summary>
        /// <param name="value">Turns underline on/off</param>
        public static implicit operator Underline?(bool value) 
            => value == false ? null 
                : new Underline {Color = ColorValue.Auto, Style = UnderlineStyle.SingleLine};

        /// <summary>
        /// Converter from Underline to bool
        /// </summary>
        /// <param name="value">Underline value</param>
        public static implicit operator bool(Underline? value)
            => value != null && value.Style != UnderlineStyle.None;

        /// <summary>
        /// Get or set the underline style for this paragraph
        /// </summary>
        public UnderlineStyle Style
        {
            get => xml.GetVal().TryGetEnumValue<UnderlineStyle>(out var result)
                ? result : UnderlineStyle.None;
            set
            {
                if (value == UnderlineStyle.None && Color.IsEmpty)
                {
                    xml?.Remove();
                    xml = null;
                    return;
                }

                Xml.SetAttributeValue(Name.MainVal, value.GetEnumName());
            }
        }

        /// <summary>
        /// Get or set the underline color for this paragraph
        /// </summary>
        public ColorValue Color
        {
            get => new(xml.AttributeValue(Name.Color),
                    xml.AttributeValue(Name.ThemeColor),
                    xml.AttributeValue(Name.ThemeTint),
                    xml.AttributeValue(Name.ThemeShade));

            set
            {
                if (Style == UnderlineStyle.None && value.IsEmpty)
                {
                    xml?.Remove();
                    xml = null;
                    return;
                }

                value.SetElementValues(Xml, Name.Color);
            }
        }

        /// <summary>
        /// Public constructor
        /// </summary>
        public Underline()
        {
            properties = null;
            xml = null;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        internal Underline(XElement properties, XElement? underlineElement)
        {
            this.properties = properties ?? throw new ArgumentNullException(nameof(properties));
            this.xml = underlineElement;
        }

    }
}
