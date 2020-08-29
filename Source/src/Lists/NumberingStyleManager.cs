using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Manager for the numbering styles (numbering.xml) in the document.
    /// </summary>
    public sealed class NumberingStyleManager : DocXBase
    {
        private readonly XDocument numberingDoc;

        /// <summary>
        /// A list of all the available numbering styles in this document.
        /// </summary>
        public IEnumerable<NumberingStyle> Styles =>
            Xml.Elements(Namespace.Main + "abstractNum").Select(e => new NumberingStyle(e));

        /// <summary>
        /// A list of all the current numbering definitions available to this document.
        /// </summary>
        public IEnumerable<NumberingDefinition> Definitions
        {
            get
            {
                var definitionsCache = Xml.Elements(Namespace.Main + "num")
                    .Select(e => new NumberingDefinition(e))
                    .ToList();
                    
                var styles = Styles.ToList();
                foreach (var item in definitionsCache)
                {
                    int id = item.StyleId;
                    item.Style = styles.Single(s => s.Id == id);
                }

                return definitionsCache.AsEnumerable();
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="documentOwner">Owning document</param>
        /// <param name="numberingPart">Numbering part</param>
        public NumberingStyleManager(IDocument documentOwner, PackagePart numberingPart) : base(documentOwner, null)
        {
            PackagePart = numberingPart ?? throw new ArgumentNullException(nameof(numberingPart));
            numberingDoc = numberingPart.Load();
            Xml = numberingDoc.Root;
        }

        /// <summary>
        /// Creates a new numbering section in the w:numbering document and adds a relationship to
        /// that section in the passed document.
        /// </summary>
        /// <param name="listType">Type of list to create</param>
        /// <param name="startNumber">Starting number</param>
        /// <returns></returns>
        public NumberingDefinition Create(NumberingFormat listType, int startNumber)
        {
            if (startNumber < 0)
                throw new ArgumentOutOfRangeException(nameof(startNumber));

            var template = listType switch
            {
                NumberingFormat.Bulleted => Resources.DefaultBulletNumberingXml(HelperFunctions.GenerateHexId()),
                NumberingFormat.Numbered => Resources.DefaultDecimalNumberingXml(HelperFunctions.GenerateHexId()),
                _ => throw new InvalidOperationException($"Unable to create {nameof(NumberingFormat)}: {listType}."),
            };

            var definitions = Definitions.ToList();
            var styles = Styles.ToList();

            int numId = definitions.Count > 0 ? definitions.Max(d => d.Id) + 1 : 1;
            int abstractNumId = styles.Count > 0 ? styles.Max(s => s.Id) + 1 : 0;

            var style = new NumberingStyle(template) {Id = abstractNumId};
            var definition = new NumberingDefinition(numId, abstractNumId) { Style = style };

            if (startNumber != 1)
            {
                definition.AddOverrideForLevel(0, new NumberingLevel { Start = startNumber });
            }

            // Style definition goes first -- this new one should be at the end of the existing styles
            var lastStyle = numberingDoc.Root!.Descendants().LastOrDefault(e => e.Name == Namespace.Main + "abstractNum");
            if (lastStyle != null)
            {
                lastStyle.AddAfterSelf(style.Xml);
            }
            // Or at the beginning of the document.
            else
            {
                numberingDoc.Root.AddFirst(style.Xml);
            }

            // Definition is always at the end of the document.
            numberingDoc.Root.Add(definition.Xml);

            return definition;
        }

        /// <summary>
        /// Save the changes back to the package.
        /// </summary>
        public void Save()
        {
            PackagePart.Save(numberingDoc);
        }

        /// <summary>
        /// Returns the starting number (with override) for the given NumId and Level
        /// </summary>
        /// <param name="numId">NumId</param>
        /// <param name="level">Level</param>
        /// <returns></returns>
        internal int GetStartingNumber(int numId, int level = 0)
        {
            var definition = Definitions.SingleOrDefault(n => n.Id == numId);
            if (definition == null)
                throw new ArgumentException("No numbering definition found.", nameof(numId));

            var levelOverride = definition.GetOverrideForLevel(level);
            return levelOverride?.Start ?? 
                   definition.Style.Levels.Single(l => l.Level == level).Start;
        }

    }
}
