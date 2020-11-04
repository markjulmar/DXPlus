using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This provides the mapping from a NumId to an abstractNumId in the /word/numbering.xml document.
    /// This mapping is needed because the Numbering Styles can be reused.
    /// </summary>
    public sealed class NumberingDefinition
    {
        internal XElement Xml { get; }

        /// <summary>
        /// Numbering Id
        /// </summary>
        public int Id => int.Parse(Xml.AttributeValue(Namespace.Main + "numId"));

        /// <summary>
        /// The abstractNum style associated to this definition.
        /// </summary>
        public int StyleId => int.Parse(Xml.Element(Namespace.Main + "abstractNumId").GetVal());

        /// <summary>
        /// Style it represents
        /// </summary>
        public NumberingStyle Style { get; internal set; }


        /// <summary>
        /// Optional override information for one or more levels.
        /// </summary>
        public IEnumerable<LevelOverride> Overrides =>
            Xml.Elements(Namespace.Main + "lvlOverride")
               .Select(xml => new LevelOverride(xml));

        /// <summary>
        /// Retrieves the level override for the given level.
        /// </summary>
        /// <param name="level">Level</param>
        /// <returns>Level override info, or null if it doesn't exist.</returns>
        public LevelOverride GetOverrideForLevel(int level) => 
            Overrides.SingleOrDefault(l => l.Level == level);

        /// <summary>
        /// Adds a new override for the given level.
        /// </summary>
        /// <param name="level">Level to override - must not already exist.</param>
        /// <param name="numberingLevel">Numbering level info</param>
        public LevelOverride AddOverrideForLevel(int level, NumberingLevel numberingLevel)
        {
            if (GetOverrideForLevel(level) != null)
                throw new ArgumentException("Level override already exists.", nameof(level));
            if (numberingLevel == null)
                throw new ArgumentNullException(nameof(numberingLevel));

            var e = new XElement(Namespace.Main + "lvlOverride",
                new XAttribute(Namespace.Main + "ilvl", level),
                numberingLevel.Xml);

            Xml.Add(e);
            return new LevelOverride(e);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="xml">Single w:num element</param>
        internal NumberingDefinition(XElement xml)
        {
            Xml = xml;
        }

        /// <summary>
        /// Constructor used to generate a new definition
        /// </summary>
        /// <param name="numId">Definition id</param>
        /// <param name="abstractNumId">Style id</param>
        internal NumberingDefinition(int numId, int abstractNumId)
        {
            Xml = new XElement(Namespace.Main + "num",
                    new XAttribute(Namespace.Main + "numId", numId),
                    new XElement(Namespace.Main + "abstractNumId",
                        new XAttribute(Name.MainVal, abstractNumId)));
        }
    }
}