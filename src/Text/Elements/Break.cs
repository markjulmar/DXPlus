using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// This class represents a page, column, or line break in a Run.
/// </summary>
public sealed class Break : TextElement
{
    /// <summary>
    /// The type of break
    /// </summary>
    public BreakType Type
    {
        get => Enum.TryParse<BreakType>(Xml.AttributeValue(Name.Type), ignoreCase:true, out var bt) ? bt : BreakType.Line;
        set => Xml!.SetAttributeValue(Name.Type, value.GetEnumName());
    }

    /// <summary>
    /// Specifies the location which shall be used as the next available line
    /// when the break type attribute has a value of line.
    /// This property only affects the restart location when the current run
    /// is being displayed on a line which does not span the full text extents
    /// due to the presence of a floating object.
    /// </summary>
    public LineBreakRestartLocation? RestartLocation
    {
        get => Type != BreakType.Line
            ? null
            : Enum.TryParse<LineBreakRestartLocation>(Xml.AttributeValue(Namespace.Main + "clear"),
                ignoreCase: true, out var bt) ? bt : null;

        set
        {
            if (Type != BreakType.Line) return;
            Xml!.SetAttributeValue(Namespace.Main + "clear", value==null ? null : value.Value.GetEnumName());
        }
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="runOwner"></param>
    /// <param name="xml"></param>
    internal Break(Run runOwner, XElement xml) : base(runOwner, xml)
    {
    }
}