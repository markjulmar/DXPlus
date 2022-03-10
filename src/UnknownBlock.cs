using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// This wraps unknown blocks found when enumerating the document.
/// </summary>
public class UnknownBlock : Block, IEquatable<UnknownBlock>
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="document"></param>
    /// <param name="packagePart"></param>
    /// <param name="xml"></param>
    internal UnknownBlock(Document? document, PackagePart? packagePart, XElement xml) : base(xml)
    {
        if (document != null)
        {
            SetOwner(document, packagePart, false);
        }
    }

    /// <summary>
    /// Returns the name of this block in the document.
    /// </summary>
    public string Name => Xml.Name.LocalName;

    /// <summary>
    /// Determines equality for an unknown block
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(UnknownBlock? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);
}