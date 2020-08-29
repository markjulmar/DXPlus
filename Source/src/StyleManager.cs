using DXPlus.Helpers;
using System;
using System.Xml.Linq;

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

    /// <summary>
    /// Specifies a set of properties which shall dictate the appearance and behavior 
    /// of a set of numbered paragraphs in a Word document. This is persisted as a w:abstractNum object
    /// in the /word/numbering.xml document.
    /// </summary>
    public sealed class NumberingStyle : DocXBase
    {
        /// <summary>
        /// Specifies a unique number which will be used as the identifier for the numbering definition.
        /// The value is referenced by numbering instances (num) via the num's abstractNumId child element.
        /// </summary>
        public int Id
        {
            get => int.Parse(Xml.AttributeValue(Namespace.Main + "abstractNumId"));
            set => Xml.SetAttributeValue(Namespace.Main + "abstractNumId", value);
        }

        /// <summary>
        /// Unique hexadecimal identifier for the numbering definition. This value will be
        /// the same for two numbering definitions based on the same initial definition (e.g.
        /// where a new definition is created from an existing one). This is persisted as w:nsid.
        /// </summary>
        public string CreatorId
        {
            get => Xml.Element(Namespace.Main + "nsid").GetVal(null);
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    Xml.Element(Namespace.Main + "nsid")?.Remove();
                }
                else
                {
                    if (!HelperFunctions.IsValidHexNumber(value))
                        throw new ArgumentException("Invalid hex value.", nameof(CreatorId));
                    Xml.AddElementVal(Namespace.Main + "nsid", value);
                }
            }
        }

        /// <summary>
        /// The type of numbering shown in the UI for lists using this definition - single, multi-level, etc.
        /// </summary>
        public NumberingLevelType LevelType
        {
            get => Xml.Element(Namespace.Main + "multiLevelType").GetVal()
                      .TryGetEnumValue<NumberingLevelType>(out var result) ? result : NumberingLevelType.None;
            set => Xml.AddElementVal(Namespace.Main + "multiLevelType",
                    value == NumberingLevelType.None ? null : value.GetEnumName());
        }

        /// <summary>
        /// Optional user-friendly name (alias) for this numbering definition.
        /// </summary>
        public string Name
        {
            get => Xml.Element(DXPlus.Name.NameId).GetVal(null);
            set => Xml.AddElementVal(DXPlus.Name.NameId, string.IsNullOrWhiteSpace(value) ? null : value);
        }

        /// <summary>
        /// Optional style name that indicates this definition doesn't contain any style properties but
        /// uses the specified style for all properties.
        /// </summary>
        public string NumStyleLink
        {
            get => Xml.Element(Namespace.Main + "numStyleLink").GetVal(null);
            set => Xml.AddElementVal(Namespace.Main + "numStyleLink", string.IsNullOrWhiteSpace(value) ? null : value);
        }

        /// <summary>
        /// Specifies this definition is the base numbering definition for the reference numbering style.
        /// </summary>
        public string StyleLink
        {
            get => Xml.Element(Namespace.Main + "styleLink").GetVal(null);
            set => Xml.AddElementVal(Namespace.Main + "styleLink", string.IsNullOrWhiteSpace(value) ? null : value);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML definition (abstractNum)</param>
        public NumberingStyle(IDocument document, XElement xml) : base(document, xml)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));
        }
    }

    public class StyleManager
    {
    }
}
