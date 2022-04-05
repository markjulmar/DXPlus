using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// A single cell in a Word table. All content in a table is contained in a cell.
/// A cell also has several properties affecting its size, appearance, and how the content it contains is formatted.
/// </summary>
public sealed class TableCell : BlockContainer, IEquatable<TableCell>
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="row">Owner Row</param>
    /// <param name="xml">XML representing this cell</param>
    internal TableCell(TableRow row, XElement xml) : base(xml)
    {
        Row = row;
        if (Row.InDocument)
        {
            SetOwner(row.Document, row.PackagePart, false);
        }
    }

    /// <summary>
    /// Row owner
    /// </summary>
    public TableRow Row { get; }

    /// <summary>
    /// Properties applied to this table cell
    /// </summary>
    public TableCellProperties Properties => new(Xml.GetOrInsertElement(Name.TableCellProperties));

    /// <summary>
    /// Gets or sets all the text for a paragraph. This will replace any existing paragraph(s)
    /// tied to the table. The <seealso cref="Paragraph"/> property can also be used to manipulate
    /// the content of the cell.
    /// </summary>
    public string Text
    {
        get => string.Join('\n', Paragraphs.Select(p => p.Text).Where(p => !string.IsNullOrEmpty(p)));
        set
        {
            switch (Paragraphs.Count())
            {
                case 0:
                    this.Add(new Paragraph(value));
                    break;
                case 1:
                    Paragraphs.First().Text = value;
                    break;
                default:
                    Xml.Elements(Name.Paragraph).Remove();
                    this.Add(new Paragraph(value));
                    break;
            }
        }
    }

    /// <summary>
    /// Determines equality for a table cell
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(TableCell? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a table cell
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as TableCell);

    /// <summary>
    /// Returns hashcode for this table
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();

}