using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents an override section in a numbering definition.
    /// </summary>
    public sealed class LevelOverride
    {
        /// <summary>
        /// The XML fragment making up this object (w:lvlOverride)
        /// </summary>
        internal XElement Xml { get; }

        /// <summary>
        /// Level to override
        /// </summary>
        public int Level
        {
            get => int.Parse(Xml.AttributeValue(Namespace.Main + "ilvl"));
            set => Xml.SetAttributeValue(Namespace.Main + "ilvl", value);
        }

        /// <summary>
        /// Specifies the number which the specified level override shall begin with.
        /// This value is used when this level initially starts in a document, as well as whenever it is restarted
        /// </summary>
        public int? Start
        {
            get => int.TryParse(Xml.Element(Namespace.Main + "startOverride").GetVal(), out var result) ? (int?)result : null;
            set => Xml.AddElementVal(Namespace.Main + "startOverride", value?.ToString());
        }

        /// <summary>
        /// Returns the details of the overriden level
        /// </summary>
        public NumberingLevel NumberingLevel => new NumberingLevel(Xml.Element(Namespace.Main + "lvl"));

        /// <summary>
        /// Removes this level override
        /// </summary>
        public void Remove()
        {
            Xml.Remove();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="xml"></param>
        internal LevelOverride(XElement xml)
        {
            Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        }
    }

    /// <summary>
    /// Represents a single level (w:lvl) in a numbering definition style.
    /// </summary>
    public sealed class NumberingLevel
    {
        /// <summary>
        /// The XML fragment making up this object (w:lvl)
        /// </summary>
        internal XElement Xml { get; }

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
        /// specifies a one-based index which determines when a numbering level
        /// should restart to its Start value. A numbering level restarts when an
        /// instance of the specified numbering level is used in the given document's contents.
        /// </summary>
        public int Restart
        {
            get => int.Parse(Xml.Element(Namespace.Main + "lvlRestart")?.GetVal() ?? "0");
            set => Xml.AddElementVal(Namespace.Main + "lvlRestart", value);
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
        /// Removes this level
        /// </summary>
        public void Remove()
        {
            Xml.Remove();
        }

        /// <summary>
        /// Public constructor for externals to add new levels/overrides.
        /// </summary>
        public NumberingLevel()
        {
            this.Xml = new XElement(Namespace.Main + "lvl");
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