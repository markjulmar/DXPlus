using DXPlus.Helpers;
using DXPlus.Resources;
using System.Diagnostics;
using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// Manager for the numbering styles (numbering.xml) in the document.
/// </summary>
public sealed class NumberingStyleManager : DocXElement
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
            var styles = Styles.ToList();
            return Xml.Elements(Namespace.Main + "num")
                .Select(e => new NumberingDefinition(e, styles))
                .ToList();
        }
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="documentOwner">Owning document</param>
    /// <param name="numberingPart">Numbering part</param>
    internal NumberingStyleManager(Document documentOwner, PackagePart numberingPart)
    {
        if (numberingPart == null)
            throw new ArgumentNullException(nameof(numberingPart));

        SetOwner(documentOwner, numberingPart, false);
        numberingDoc = numberingPart.Load();
        Xml = numberingDoc.Root ?? throw new DocumentFormatException("NumberingDoc");
    }

    /// <summary>
    /// Creates a new numbering section in the w:numbering document and adds a relationship to
    /// that section in the passed document.
    /// </summary>
    /// <param name="listType">Type of list to create</param>
    /// <param name="startNumber">Starting number</param>
    /// <returns></returns>
    public NumberingDefinition Create(NumberingFormat listType, int startNumber = 1)
    {
        Debug.Assert(numberingDoc != null);

        if (startNumber < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(startNumber));
        }

        var styles = Styles.ToList();
        var definitions = Definitions.ToList();

        // TODO: improve the search.
        NumberingStyle? style = null; // styles.FirstOrDefault(s => s.Levels.FirstOrDefault()?.Format == listType);
        if (style == null)
        {
            var template = listType switch
            {
                NumberingFormat.Bullet => Resource.DefaultBulletNumberingXml(HelperFunctions.GenerateHexId()),
                NumberingFormat.Numbered => Resource.DefaultDecimalNumberingXml(HelperFunctions.GenerateHexId()),
                _ => throw new InvalidOperationException($"Unable to create {nameof(NumberingFormat)}: {listType}."),
            };
            int abstractNumId = styles.Count > 0 ? styles.Max(s => s.Id) + 1 : 0;
            style = new NumberingStyle(template) { Id = abstractNumId };

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
        }

        int numId = definitions.Count > 0 ? definitions.Max(d => d.Id) + 1 : 1;
        var definition = new NumberingDefinition(numId, style);

        if (startNumber != 1)
        {
            definition.AddOverrideForLevel(0, new NumberingLevel { Start = startNumber });
        }

        // Definition is always at the end of the document.
        numberingDoc.Root!.Add(definition.Xml);

        return definition;
    }

    /// <summary>
    /// Save the changes back to the package.
    /// </summary>
    internal void Save()
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