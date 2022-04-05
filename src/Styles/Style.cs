using System.Diagnostics;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Style definition which groups all the style properties.
/// </summary>
[DebuggerDisplay("{Id} - {Name} Type={Type} Default={IsDefault}")]
public sealed class Style : XElementWrapper
{
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// Unique id for this style
    /// </summary>
    public string? Id
    {
        get => Xml.AttributeValue(Namespace.Main + "styleId");
        set => Xml.SetAttributeValue(Namespace.Main + "styleId", value);
    }

    /// <summary>
    /// Name of the style
    /// </summary>
    public string? Name
    {
        get => Xml.Element(Internal.Name.NameId).GetVal();
        set => Xml.AddElementVal(Internal.Name.NameId, value);
    }

    /// <summary>
    /// Style is a user-defined style.
    /// </summary>
    public bool IsCustom
    {
        get => Xml.BoolAttributeValue(Namespace.Main + "customStyle") == true;
        set => Xml.SetAttributeValue(Namespace.Main + "customStyle", value ? "1" : null);
    }

    /// <summary>
    /// Specifies that this style is the default for the given Type.
    /// </summary>
    public bool IsDefault
    {
        get => Xml.BoolAttributeValue(Namespace.Main + "default") == true;
        set => Xml.SetAttributeValue(Namespace.Main + "default", value ? "1" : null);
    }

    /// <summary>
    /// The type this style is applied to
    /// </summary>
    public StyleType Type
    {
        get => Xml.AttributeValue(Namespace.Main + "type").TryGetEnumValue<StyleType>(out var result)
            ? result
            : StyleType.Paragraph;

        set => Xml.SetAttributeValue(Namespace.Main + "type", value.GetEnumName());
    }

    /// <summary>
    /// Returns all the optional table styles assigned to this style.
    /// </summary>
    public IList<TableStyle> TableStyles { get; }

    /// <summary>
    /// The optional paragraph properties to apply to this table section.
    /// </summary>
    public ParagraphProperties? ParagraphFormatting
    {
        get
        {
            var xml = Xml.Element(Internal.Name.ParagraphProperties);
            return xml != null ? new ParagraphProperties(xml) : null;
        }

        set
        {
            Xml.Element(Internal.Name.ParagraphProperties)?.Remove();
            if (value == null) return;

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();
            Xml.Add(xml);
        }
    }

    /// <summary>
    /// The optional run properties to apply to this table section.
    /// </summary>
    public Formatting? Formatting
    {
        get
        {
            var xml = Xml.Element(Internal.Name.RunProperties);
            return xml != null ? new Formatting(xml) : null;
        }

        set
        {
            Xml.Element(Internal.Name.RunProperties)?.Remove();
            if (value == null) return;

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();
            Xml.Add(xml);
        }
    }

    /// <summary>
    /// The optional table properties to apply to this table section.
    /// </summary>
    public TableProperties? TableFormatting
    {
        get
        {
            var xml = Xml.Element(Internal.Name.TableProperties);
            return xml != null ? new TableProperties(xml, null) : null;
        }

        set
        {
            Xml.Element(Internal.Name.TableProperties)?.Remove();
            if (value == null) return;

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();
            Xml.Add(xml);
        }
    }

    /// <summary>
    /// The optional table row properties to apply to this table section.
    /// </summary>
    public TableRowProperties? TableRowFormatting
    {
        get
        {
            var xml = Xml.Element(Internal.Name.TableRowProperties);
            return xml != null ? new TableRowProperties(xml) : null;
        }

        set
        {
            Xml.Element(Internal.Name.TableRowProperties)?.Remove();
            if (value == null) return;

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();
            Xml.Add(xml);
        }
    }

    /// <summary>
    /// The optional table cell properties to apply to this table section.
    /// </summary>
    public TableCellProperties? TableCellFormatting
    {
        get
        {
            var xml = Xml.Element(Internal.Name.TableCellProperties);
            return xml != null ? new TableCellProperties(xml) : null;
        }

        set
        {
            Xml.Element(Internal.Name.TableCellProperties)?.Remove();
            if (value == null) return;

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();
            Xml.Add(xml);
        }
    }

    /// <summary>
    /// The style this one is based on.
    /// </summary>
    public string? BasedOn
    {
        get => Xml.Element(Namespace.Main + "basedOn").GetVal(null);
        set => Xml.AddElementVal(Namespace.Main + "basedOn", string.IsNullOrWhiteSpace(value) ? null : value);
    }

    /// <summary>
    /// The default style for the next paragraph
    /// </summary>
    public string? NextParagraphStyle
    {
        get => Xml.Element(Namespace.Main + "next").GetVal(null);
        set => Xml.AddElementVal(Namespace.Main + "next", string.IsNullOrWhiteSpace(value) ? null : value);
    }

    /// <summary>
    /// Linked style
    /// </summary>
    public string? LinkedStyle
    {
        get => Xml.Element(Namespace.Main + "link").GetVal(null);
        set => Xml.AddElementVal(Namespace.Main + "link", string.IsNullOrWhiteSpace(value) ? null : value);
    }

    // TODO: add tblPr, tblStylePr, tcPr, trPr

    /// <summary>
    /// Constructor for an existing style
    /// </summary>
    /// <param name="xml">Element in the style document</param>
    internal Style(XElement xml)
    {
        base.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        TableStyles = new TableStyleCollection(Xml);
    }

    /// <summary>
    /// Constructor to add a new style to the document.
    /// </summary>
    /// <param name="owner">Style manager owner</param>
    /// <param name="id">ID for the style</param>
    /// <param name="name">Name of the style</param>
    /// <param name="type">Style type</param>
    /// <param name="latentStyle"></param>
    internal Style(XDocument owner, string id, string name, StyleType type, XElement? latentStyle)
    {
        if (owner == null) throw new ArgumentNullException(nameof(owner));

        /* 	<w:style w:customStyle="1" w:styleId="Normal" w:type="paragraph">
                <w:name w:val="Normal"/>
                <w:qFormat/>
            </w:style>  */

        base.Xml = new XElement(Namespace.Main + "style",
            new XAttribute(Namespace.Main + "styleId", id),
            new XAttribute(Namespace.Main + "type", type.ToString().ToLower()));

        if (latentStyle == null)
        {
            Xml.Add(new XElement(Namespace.Main + "name", new XAttribute(Namespace.Main + "val", name)));
            Xml.Add(new XAttribute(Namespace.Main + "customStyle", 1));
            Xml.Add(new XElement(Namespace.Main + "qFormat"));
        }
        else
        {
            //<w:lsdException w:name="caption" w:qFormat="1" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1"/>
            foreach (var attr in latentStyle.Attributes())
            {
                if (attr.Value == "1" || attr.Value.ToLower() == "true")
                {
                    Xml.Add(new XElement(attr.Name));
                }
                else if (attr.Value != "0")
                {
                    Xml.Add(new XElement(attr.Name, new XAttribute(Namespace.Main + "val", attr.Value)));
                }
            }
        }

        owner.Root!.Add(Xml);
        TableStyles = new TableStyleCollection(Xml);
    }
}