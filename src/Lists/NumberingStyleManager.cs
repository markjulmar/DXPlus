using DXPlus.Resources;
using System.Diagnostics;
using System.Drawing;
using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

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
    /// <returns></returns>
    public NumberingDefinition BulletStyle()
    {
        Debug.Assert(numberingDoc != null);

        var styles = Styles.ToList();
        var style = styles.FirstOrDefault(s => s.Levels.Any(nl => nl.Level == 0 && nl.Format == NumberingFormat.Bullet))
            ?? AddNumberingStyle(Resource.DefaultBulletNumberingXml(DocumentHelpers.GenerateHexId()), styles);
        
        return AddNumberingDefinition(1, style);
    }

    /// <summary>
    /// Creates a new numbering section in the w:numbering document and adds a relationship to
    /// that section in the passed document.
    /// </summary>
    /// <returns></returns>
    public NumberingDefinition NumberStyle(int? startNumber = null)
    {
        Debug.Assert(numberingDoc != null);
        if (startNumber < 1) throw new ArgumentOutOfRangeException(nameof(startNumber));

        var styles = Styles.ToList();
        var style = styles.FirstOrDefault(s => s.Levels.Any(nl => nl.Level == 0 && nl.Format == NumberingFormat.Numbered))
            ?? AddNumberingStyle(Resource.DefaultBulletNumberingXml(DocumentHelpers.GenerateHexId()), styles);

        return AddNumberingDefinition(startNumber??1, style);
    }

    /// <summary>
    /// Creates a new numbering section in the w:numbering document and adds a relationship to
    /// that section in the passed document.
    /// </summary>
    /// <param name="text">Text to use for 1st level of list.</param>
    /// <param name="fontFamily">Font to use for text, defaults to Courier New</param>
    /// <returns></returns>
    public NumberingDefinition CustomBulletStyle(string text, FontFamily? fontFamily = null)
    {
        Debug.Assert(numberingDoc != null);

        var styles = Styles.ToList();
        var style = styles.FirstOrDefault(s => s.Levels.Any(nl => nl.Level == 0 && nl.Text == text)) 
            ?? AddNumberingStyle(Resource.CustomBulletNumberingXml(DocumentHelpers.GenerateHexId(), 
                    text, fontFamily?.Name), styles);

        return AddNumberingDefinition(1, style);
    }

    /// <summary>
    /// Method to create a new {abstractNum} style definition from a passed XML template.
    /// </summary>
    /// <param name="styleTemplate">XML template</param>
    /// <param name="styles">Existing style list.</param>
    /// <returns>New added numbering style</returns>
    private NumberingStyle AddNumberingStyle(XElement styleTemplate, IReadOnlyCollection<NumberingStyle> styles)
    {
        int abstractNumId = styles.Count > 0 ? styles.Max(s => s.Id) + 1 : 0;
        var style = new NumberingStyle(styleTemplate) { Id = abstractNumId };

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

        return style;

    }

    /// <summary>
    /// Method to create a new {w:num} tied to an abstract definition in our numbering document.
    /// </summary>
    /// <param name="startNumber">Starting number</param>
    /// <param name="style"></param>
    /// <returns></returns>
    private NumberingDefinition AddNumberingDefinition(int startNumber, NumberingStyle style)
    {
        var definitions = Definitions.ToList();
        int numId = definitions.Count > 0 ? definitions.Max(d => d.Id) + 1 : 1;
        var definition = new NumberingDefinition(numId, style);

        if (startNumber != 1)
            definition.AddOverrideForLevel(0, new NumberingLevel {Start = startNumber});

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