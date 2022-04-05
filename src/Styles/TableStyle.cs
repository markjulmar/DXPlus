using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Specifies a set of formatting properties which shall be conditionally applied to the
/// parts of a table which match the requirement specified on the Type property.
/// </summary>
public sealed class TableStyle : XElementWrapper
{
    /// <summary>
    /// The XML representing the table style.
    /// </summary>
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// Specifies the section of the table to which the formatting properties should be applied. 
    /// </summary>
    public TableStyleType Type
    {
        get => Xml.AttributeValue(Namespace.Main + "type").TryGetEnumValue<TableStyleType>(out var result)
                ? result
                : TableStyleType.WholeTable;

        set => Xml.SetAttributeValue(Namespace.Main + "type", value.GetEnumName());
    }

    /// <summary>
    /// The optional paragraph properties to apply to this table section.
    /// </summary>
    public ParagraphProperties? ParagraphFormatting
    {
        get
        {
            var xml = Xml.Element(Name.ParagraphProperties);
            return xml != null ? new ParagraphProperties(xml) : null;
        }

        set
        {
            Xml.Element(Name.ParagraphProperties)?.Remove();
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
            var xml = Xml.Element(Name.RunProperties);
            return xml != null ? new Formatting(xml) : null;
        }

        set
        {
            Xml.Element(Name.RunProperties)?.Remove();
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
            var xml = Xml.Element(Name.TableProperties);
            return xml != null ? new TableProperties(xml, null) : null;
        }

        set
        {
            Xml.Element(Name.TableProperties)?.Remove();
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
            var xml = Xml.Element(Name.TableRowProperties);
            return xml != null ? new TableRowProperties(xml) : null;
        }

        set
        {
            Xml.Element(Name.TableRowProperties)?.Remove();
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
            var xml = Xml.Element(Name.TableCellProperties);
            return xml != null ? new TableCellProperties(xml) : null;
        }

        set
        {
            Xml.Element(Name.TableCellProperties)?.Remove();
            if (value == null) return;

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();
            Xml.Add(xml);
        }
    }

    /// <summary>
    /// Add a new table style
    /// </summary>
    public TableStyle(TableStyleType type) : this(new XElement(Name.TableStyles))
    {
        Type = type;
    }

    /// <summary>
    /// Add a new table style
    /// </summary>
    /// <param name="element">XML element to wrap</param>
    internal TableStyle(XElement element)
    {
        base.Xml = element;
    }
}