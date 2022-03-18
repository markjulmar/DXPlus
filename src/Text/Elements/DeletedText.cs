using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// Wraps a delText entry under a Run
/// </summary>
public sealed class DeletedText : Text
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="runOwner"></param>
    /// <param name="xml"></param>
    internal DeletedText(Run runOwner, XElement xml) : base(runOwner, xml)
    {
    }
}