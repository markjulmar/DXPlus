using System.Collections;
using System.ComponentModel;
using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Manager for the named styles (styles.xml) in the document.
/// </summary>
public sealed class StyleManager : DocXElement, IEnumerable<Style>
{
    private readonly XDocument stylesDoc;

    /// <summary>
    /// Get all the latent styles from the document
    /// </summary>
    public IEnumerable<string> LatentStyles =>
        Xml.Elements(Namespace.Main + "lsdException")
            .Select(e => e.Attribute(Namespace.Main + "name")?.Value)
            .OmitNull();

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="documentOwner">Owning document</param>
    /// <param name="stylesPart">Numbering part</param>
    internal StyleManager(Document documentOwner, PackagePart stylesPart)
    {
        if (documentOwner == null) throw new ArgumentNullException(nameof(documentOwner));
        if (stylesPart == null) throw new ArgumentNullException(nameof(stylesPart));

        SetOwner(documentOwner, stylesPart, false);

        stylesDoc = stylesPart.Load();
        Xml = stylesDoc.Root ?? throw new DocumentFormatException(nameof(StyleManager));
    }

    /// <summary>
    /// Save the changes back to the package.
    /// </summary>
    internal void Save()
    {
        PackagePart.Save(stylesDoc);
    }

    /// <summary>
    /// Returns whether the given style exists in the style catalog
    /// </summary>
    /// <param name="styleId"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    public bool Exists(string styleId, StyleType type)
    {
        return stylesDoc.Descendants(Namespace.Main + "style").Any(x =>
            x.AttributeValue(Namespace.Main + "type")?.Equals(type.GetEnumName()) == true
            && x.AttributeValue(Namespace.Main + "styleId")?.Equals(styleId) == true);
    }

    /// <summary>
    /// This method adds a new style to the document.
    /// </summary>
    /// <param name="name">Name of the style</param>
    /// <param name="id">Style identifier</param>
    /// <param name="type">Style type</param>
    /// <param name="basedOnStyle"></param>
    /// <returns>Created style which can be edited</returns>
    public Style Add(string id, string name, StyleType type, Style? basedOnStyle = null)
    {
        if (string.IsNullOrWhiteSpace(id))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(id));
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
        if (!Enum.IsDefined(typeof(StyleType), type))
            throw new InvalidEnumArgumentException(nameof(type), (int) type, typeof(StyleType));

        basedOnStyle ??= this.FindDefault(type);
        id = new string(id.Where(char.IsLetterOrDigit).ToArray());

        // If the style is a default one, pick off the exception data.
        var lsdException = Xml.Descendants(Namespace.Main + "lsdException")
            .SingleOrDefault(x => string.Compare(x.Attribute(Namespace.Main + "name")?.Value, name, StringComparison.InvariantCultureIgnoreCase) == 0);

        return new Style(stylesDoc, id, name, type, lsdException) { BasedOn = basedOnStyle?.Id };
    }

    /// <summary>
    /// Returns the default style for a given type.
    /// </summary>
    /// <param name="type">Type to search for</param>
    /// <returns>Style if it exists, null if not.</returns>
    public Style? FindDefault(StyleType type) => this.FirstOrDefault(s => s.IsDefault && s.Type == type);

    /// <summary>
    /// This method retrieves the XML block associated with a style.
    /// </summary>
    /// <param name="styleId">Id</param>
    /// <param name="type">Style type</param>
    /// <returns>Style if present</returns>
    public Style? Find(string styleId, StyleType type) =>
        this.SingleOrDefault(s => s.Id == styleId && s.Type == type);

    /// <summary>
    /// This method adds a new Style XML block to the /word/styles.xml document
    /// </summary>
    /// <param name="xml">XML to add</param>
    internal void AddStyle(XElement xml)
    {
        if (xml == null)
            throw new ArgumentNullException(nameof(xml));

        if (xml.Name != Namespace.Main + "style")
            throw new ArgumentException("Passed element is not a <style> object.", nameof(xml));

        stylesDoc.Root!.Add(xml);
    }

    /// <summary>Returns an enumerator that iterates through the collection.</summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<Style> GetEnumerator() => 
        Xml.Elements(Namespace.Main + "style")
            .Select(e => new Style(e))
            .GetEnumerator();

    /// <summary>Returns an enumerator that iterates through a collection.</summary>
    /// <returns>An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}