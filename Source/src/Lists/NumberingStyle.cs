using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Specifies a set of properties which shall dictate the appearance and behavior 
    /// of a set of numbered paragraphs in a Word document. This is persisted as a w:abstractNum object
    /// in the /word/numbering.xml document.
    /// </summary>
    public sealed class NumberingStyle
    {
        internal XElement Xml { get; }

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
        /// The levels
        /// </summary>
        public IEnumerable<NumberingLevel> Levels
        {
            get => Xml.Elements(Namespace.Main + "lvl")
                    .Select(e => new NumberingLevel(e))
                    .OrderBy(nl => nl.Level);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="xml">XML definition (abstractNum)</param>
        public NumberingStyle(XElement xml)
        {
            Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        }
    }
}