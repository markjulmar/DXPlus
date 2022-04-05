using System.Collections;
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
public sealed class NumberingStyleManager : DocXElement, IReadOnlyList<NumberingDefinition>
{
    private readonly XDocument numberingDoc;
    private readonly XElementCollection<NumberingStyle> numberingStyles;
    private readonly XElementCollection<NumberingDefinition> numberingDefinitions;

    /// <summary>
    /// Returns the defined abstract styles.
    /// </summary>
    public IReadOnlyList<NumberingStyle> NumberingStyles => numberingStyles;

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

        numberingStyles =
            new XElementCollection<NumberingStyle>(Xml, null, Namespace.Main + "abstractNum",
                e => new NumberingStyle(e), isReadOnly: true);

        numberingDefinitions =
            new XElementCollection<NumberingDefinition>(Xml, null, Namespace.Main + "num",
                e => new NumberingDefinition(e, numberingStyles), isReadOnly: true);
    }

    /// <summary>
    /// Creates a new circular bullet style and related numbering definition and adds it to the document.
    /// The return style on the returned NumberingDefinition can be changed.
    /// </summary>
    /// <returns></returns>
    public NumberingDefinition AddBulletDefinition()
    {
        Debug.Assert(numberingDoc != null);

        var ns = new NumberingStyle(Resource.DefaultBulletNumberingXml(DocumentHelpers.GenerateHexId()));
        AddNumberingStyle(ns);
        return CreateNumberingDefinition(ns);
    }

    /// <summary>
    /// Creates a new numbered (1.2.3) style and related numbering definition and adds it to the document
    /// The return style on the returned NumberingDefinition can be changed.
    /// </summary>
    /// <returns></returns>
    public NumberingDefinition AddNumberedDefinition(int startNumber)
    {
        Debug.Assert(numberingDoc != null);
        if (startNumber < 1) throw new ArgumentOutOfRangeException(nameof(startNumber));

        var ns = new NumberingStyle(Resource.DefaultDecimalNumberingXml(DocumentHelpers.GenerateHexId()));
        AddNumberingStyle(ns);
        return CreateNumberingDefinition(ns, startNumber);
    }

    /// <summary>
    /// Creates a new custom bullet style with a single level and related numbering definition and
    /// adds it to the document. The return style on the returned NumberingDefinition can be changed.
    /// </summary>
    /// <param name="text">Text to use for 1st level of list.</param>
    /// <param name="fontFamily">Font to use for text, defaults to Courier New</param>
    /// <returns></returns>
    public NumberingDefinition AddCustomDefinition(string text, FontFamily? fontFamily = null)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));
        Debug.Assert(numberingDoc != null);

        var ns = new NumberingStyle(Resource.CustomBulletNumberingXml(DocumentHelpers.GenerateHexId(), text, fontFamily?.Name));
        AddNumberingStyle(ns);
        return CreateNumberingDefinition(ns);
    }

    /// <summary>
    /// Get the next available abstract id.
    /// </summary>
    private int NextAbstractId => this.Any() ? this.Max(s => s.Id) + 1 : 0;

    /// <summary>
    /// Adds a new numbering style to the available styles.
    /// </summary>
    /// <param name="style">Defined abstract numbering style</param>
    public void AddNumberingStyle(NumberingStyle style)
    {
        if (style == null) throw new ArgumentNullException(nameof(style));
        if (style.Id != -1) throw new ArgumentException("Style already added to document.", nameof(style));

        style.Id = NextAbstractId;

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

    /// <summary>
    /// Method to create a new usable numbering definition from an existing style.
    /// </summary>
    /// <param name="style">Created style</param>
    /// <returns>Numbering definition that can be applied to a list.</returns>
    public NumberingDefinition CreateNumberingDefinition(NumberingStyle style) => CreateNumberingDefinition(style, 1);

    /// <summary>
    /// Method to create a new usable numbering definition from an existing style.
    /// </summary>
    /// <param name="style">Created style</param>
    /// <param name="startNumber">Starting number</param>
    /// <returns>Numbering definition that can be applied to a list.</returns>
    public NumberingDefinition CreateNumberingDefinition(NumberingStyle style, int startNumber)
    {
        if (style == null) throw new ArgumentNullException(nameof(style));
        if (style.Id == -1)
        {
            AddNumberingStyle(style);
        }

        var definitions = this.ToList();
        int numId = definitions.Count > 0 ? definitions.Max(d => d.Id) + 1 : 1;
        var definition = new NumberingDefinition(numId, style);

        if (startNumber != 1)
        {
            definition.AddOverrideForLevel(0)
                .NumberingLevel.Start = startNumber;
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

    /// <summary>Returns an enumerator that iterates through the collection.</summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<NumberingDefinition> GetEnumerator() => numberingDefinitions.GetEnumerator();

    /// <summary>Returns an enumerator that iterates through a collection.</summary>
    /// <returns>An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Gets the number of elements in the collection.</summary>
    /// <returns>The number of elements in the collection.</returns>
    public int Count => numberingDefinitions.Count;

    /// <summary>Gets the element at the specified index in the read-only list.</summary>
    /// <param name="index">The zero-based index of the element to get.</param>
    /// <returns>The element at the specified index in the read-only list.</returns>
    public NumberingDefinition this[int index] => numberingDefinitions[index];
}