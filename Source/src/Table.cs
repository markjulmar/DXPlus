using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{

    /// <summary>
    /// Represents a Table in a document.
    /// </summary>
    public class Table : InsertBeforeOrAfter
    {
        private string customTableDesignName;
        private string tableCaption;
        private string tableDescription;
        private TableDesign tableDesign;
        private Alignment alignment;
        private AutoFit autoFit;
        private int? cachedColCount;
        private float[] ColumnWidthsValue;

        internal Table(DocX document, XElement xml) : base(document, xml)
        {
            autoFit = AutoFit.ColumnWidth;
            Xml = xml;
            packagePart = document.packagePart;

            XElement properties = xml.Element(DocxNamespace.Main + "tblPr");
            XElement style = properties.Element(DocxNamespace.Main + "tblStyle");
            if (style != null)
            {
                XAttribute val = style.Attribute(DocxNamespace.Main + "val");
                tableDesign = val != null
                    ? Enum.TryParse<TableDesign>(val.Value.Replace("-", ""), out var result) ? result : TableDesign.Custom
                    : TableDesign.None;
            }
            else
            {
                tableDesign = TableDesign.None;
            }

            XElement tableLook = properties.Element(DocxNamespace.Main + "tblLook");
            if (tableLook != null)
            {
                TableLook = new TableLook
                {
                    FirstRow = tableLook.AttributeValue(DocxNamespace.Main + "firstRow") == "1",
                    LastRow = tableLook.AttributeValue(DocxNamespace.Main + "lastRow") == "1",
                    FirstColumn = tableLook.AttributeValue(DocxNamespace.Main + "firstColumn") == "1",
                    LastColumn = tableLook.AttributeValue(DocxNamespace.Main + "lastColumn") == "1",
                    NoHorizontalBanding = tableLook.AttributeValue(DocxNamespace.Main + "noHBand") == "1",
                    NoVerticalBanding = tableLook.AttributeValue(DocxNamespace.Main + "noVBand") == "1"
                };
            }
        }

        public Alignment Alignment
        {
            get => alignment;
            set
            {
                alignment = value;

                XElement tblPr = Xml.Descendants(DocxNamespace.Main + "tblPr").First();
                tblPr.Descendants(DocxNamespace.Main + "jc").FirstOrDefault()?.Remove();
                tblPr.Add(new XElement(DocxNamespace.Main + "jc",
                        new XAttribute(DocxNamespace.Main + "val",
                            value.ToString().ToLower())));
            }
        }

        /// <summary>
        /// Auto size this table according to some rule.
        /// </summary>
        public AutoFit AutoFit
        {
            get => autoFit;

            set
            {
                string tableAttributeValue = string.Empty;
                string columnAttributeValue = string.Empty;
                switch (value)
                {
                    case AutoFit.ColumnWidth:
                        {
                            tableAttributeValue = "auto";
                            columnAttributeValue = "dxa";

                            // Disable "Automatically resize to fit contents" option
                            XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
                            if (tblPr != null)
                            {
                                XElement layout = tblPr.Element(DocxNamespace.Main + "tblLayout");
                                if (layout == null)
                                {
                                    tblPr.Add(new XElement(DocxNamespace.Main + "tblLayout"));
                                    layout = tblPr.Element(DocxNamespace.Main + "tblLayout");
                                }

                                XAttribute type = layout.Attribute(DocxNamespace.Main + "type");
                                if (type == null)
                                {
                                    layout.Add(new XAttribute(DocxNamespace.Main + "type", String.Empty));
                                    type = layout.Attribute(DocxNamespace.Main + "type");
                                }

                                type.Value = "fixed";
                            }

                            break;
                        }

                    case AutoFit.Contents:
                        tableAttributeValue = columnAttributeValue = "auto";
                        break;

                    case AutoFit.Window:
                        tableAttributeValue = columnAttributeValue = "pct";
                        break;

                    case AutoFit.Fixed:
                        {
                            tableAttributeValue = columnAttributeValue = "dxa";

                            XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
                            XElement tblLayout = tblPr.Element(DocxNamespace.Main + "tblLayout");

                            if (tblLayout == null)
                            {
                                XElement tmp = tblPr.Element(DocxNamespace.Main + "tblInd");
                                if (tmp == null)
                                {
                                    tmp = tblPr.Element(DocxNamespace.Main + "tblW");
                                }

                                tmp.AddAfterSelf(new XElement(DocxNamespace.Main + "tblLayout"));
                                tmp = tblPr.Element(DocxNamespace.Main + "tblLayout");
                                tmp.SetAttributeValue(DocxNamespace.Main + "type", "fixed");

                                tmp = tblPr.Element(DocxNamespace.Main + "tblW");
                                Double i = 0;
                                foreach (Double w in ColumnWidths)
                                    i += w;

                                tmp.SetAttributeValue(DocxNamespace.Main + "w", i.ToString());

                                break;
                            }
                            else
                            {
                                var qry = from d in Xml.Descendants()
                                          let type = d.Attribute(DocxNamespace.Main + "type")
                                          where (d.Name.LocalName == "tblLayout") && type != null
                                          select type;

                                foreach (XAttribute type in qry)
                                    type.Value = "fixed";

                                XElement tmp = tblPr.Element(DocxNamespace.Main + "tblW");
                                Double i = 0;
                                foreach (Double w in ColumnWidths)
                                    i += w;

                                tmp.SetAttributeValue(DocxNamespace.Main + "w", i.ToString());
                                break;
                            }
                        }
                }

                // Set table attributes
                var query = from d in Xml.Descendants()
                            let type = d.Attribute(DocxNamespace.Main + "type")
                            where (d.Name.LocalName == "tblW") && type != null
                            select type;

                foreach (XAttribute type in query)
                    type.Value = tableAttributeValue;

                // Set column attributes
                query = from d in Xml.Descendants()
                        let type = d.Attribute(DocxNamespace.Main + "type")
                        where (d.Name.LocalName == "tcW") && type != null
                        select type;

                foreach (XAttribute type in query)
                    type.Value = columnAttributeValue;

                autoFit = value;
            }
        }

        /// <summary>
        /// Returns the number of columns in this table.
        /// </summary>
        public int ColumnCount => cachedColCount ?? (RowCount == 0 ? 0 : (cachedColCount = Rows[0].ColumnCount).Value);

        /// <summary>
        /// Gets a list of all column widths for this table.
        /// </summary>
        public List<double> ColumnWidths
        {
            get
            {
                // Get the table grid properties
                XElement grid = Xml.Element(DocxNamespace.Main + "tblGrid");
                if (grid == null)
                    return null;

                // get grid column values
                var cols = grid.Elements(DocxNamespace.Main + "gridCol");
                if (cols == null)
                    return null;

                // Convert to doubles
                return cols.Select(c => Convert.ToDouble(c.AttributeValue(DocxNamespace.Main + "w")))
                    .ToList();
            }
        }

        /// <summary>
        /// Extra property for Custom Table Style provided by carpfisher - Thanks
        /// </summary>
        public string CustomTableDesignName
        {
            set
            {
                customTableDesignName = value;
                this.Design = TableDesign.Custom;
            }

            get
            {
                return customTableDesignName;
            }
        }

        /// <summary>
        /// The design\style to apply to this table.
        /// </summary>
        public TableDesign Design
        {
            get => tableDesign;
            set
            {
                XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
                XElement style = tblPr.Element(DocxNamespace.Main + "tblStyle");
                if (style == null)
                {
                    tblPr.Add(new XElement(DocxNamespace.Main + "tblStyle"));
                    style = tblPr.Element(DocxNamespace.Main + "tblStyle");
                }

                XAttribute val = style.Attribute(DocxNamespace.Main + "val");
                if (val == null)
                {
                    style.Add(new XAttribute(DocxNamespace.Main + "val", ""));
                    val = style.Attribute(DocxNamespace.Main + "val");
                }

                tableDesign = value;

                if (tableDesign == TableDesign.None)
                {
                    style?.Remove();
                }

                if (tableDesign == TableDesign.Custom)
                {
                    if (string.IsNullOrEmpty(customTableDesignName))
                    {
                        tableDesign = TableDesign.None;
                        style?.Remove();
                    }
                    else
                    {
                        val.Value = customTableDesignName;
                    }
                }
                else
                {
                    switch (tableDesign)
                    {
                        case TableDesign.TableNormal:
                            val.Value = "TableNormal";
                            break;

                        case TableDesign.TableGrid:
                            val.Value = "TableGrid";
                            break;

                        case TableDesign.LightShading:
                            val.Value = "LightShading";
                            break;

                        case TableDesign.LightShadingAccent1:
                            val.Value = "LightShading-Accent1";
                            break;

                        case TableDesign.LightShadingAccent2:
                            val.Value = "LightShading-Accent2";
                            break;

                        case TableDesign.LightShadingAccent3:
                            val.Value = "LightShading-Accent3";
                            break;

                        case TableDesign.LightShadingAccent4:
                            val.Value = "LightShading-Accent4";
                            break;

                        case TableDesign.LightShadingAccent5:
                            val.Value = "LightShading-Accent5";
                            break;

                        case TableDesign.LightShadingAccent6:
                            val.Value = "LightShading-Accent6";
                            break;

                        case TableDesign.LightList:
                            val.Value = "LightList";
                            break;

                        case TableDesign.LightListAccent1:
                            val.Value = "LightList-Accent1";
                            break;

                        case TableDesign.LightListAccent2:
                            val.Value = "LightList-Accent2";
                            break;

                        case TableDesign.LightListAccent3:
                            val.Value = "LightList-Accent3";
                            break;

                        case TableDesign.LightListAccent4:
                            val.Value = "LightList-Accent4";
                            break;

                        case TableDesign.LightListAccent5:
                            val.Value = "LightList-Accent5";
                            break;

                        case TableDesign.LightListAccent6:
                            val.Value = "LightList-Accent6";
                            break;

                        case TableDesign.LightGrid:
                            val.Value = "LightGrid";
                            break;

                        case TableDesign.LightGridAccent1:
                            val.Value = "LightGrid-Accent1";
                            break;

                        case TableDesign.LightGridAccent2:
                            val.Value = "LightGrid-Accent2";
                            break;

                        case TableDesign.LightGridAccent3:
                            val.Value = "LightGrid-Accent3";
                            break;

                        case TableDesign.LightGridAccent4:
                            val.Value = "LightGrid-Accent4";
                            break;

                        case TableDesign.LightGridAccent5:
                            val.Value = "LightGrid-Accent5";
                            break;

                        case TableDesign.LightGridAccent6:
                            val.Value = "LightGrid-Accent6";
                            break;

                        case TableDesign.MediumShading1:
                            val.Value = "MediumShading1";
                            break;

                        case TableDesign.MediumShading1Accent1:
                            val.Value = "MediumShading1-Accent1";
                            break;

                        case TableDesign.MediumShading1Accent2:
                            val.Value = "MediumShading1-Accent2";
                            break;

                        case TableDesign.MediumShading1Accent3:
                            val.Value = "MediumShading1-Accent3";
                            break;

                        case TableDesign.MediumShading1Accent4:
                            val.Value = "MediumShading1-Accent4";
                            break;

                        case TableDesign.MediumShading1Accent5:
                            val.Value = "MediumShading1-Accent5";
                            break;

                        case TableDesign.MediumShading1Accent6:
                            val.Value = "MediumShading1-Accent6";
                            break;

                        case TableDesign.MediumShading2:
                            val.Value = "MediumShading2";
                            break;

                        case TableDesign.MediumShading2Accent1:
                            val.Value = "MediumShading2-Accent1";
                            break;

                        case TableDesign.MediumShading2Accent2:
                            val.Value = "MediumShading2-Accent2";
                            break;

                        case TableDesign.MediumShading2Accent3:
                            val.Value = "MediumShading2-Accent3";
                            break;

                        case TableDesign.MediumShading2Accent4:
                            val.Value = "MediumShading2-Accent4";
                            break;

                        case TableDesign.MediumShading2Accent5:
                            val.Value = "MediumShading2-Accent5";
                            break;

                        case TableDesign.MediumShading2Accent6:
                            val.Value = "MediumShading2-Accent6";
                            break;

                        case TableDesign.MediumList1:
                            val.Value = "MediumList1";
                            break;

                        case TableDesign.MediumList1Accent1:
                            val.Value = "MediumList1-Accent1";
                            break;

                        case TableDesign.MediumList1Accent2:
                            val.Value = "MediumList1-Accent2";
                            break;

                        case TableDesign.MediumList1Accent3:
                            val.Value = "MediumList1-Accent3";
                            break;

                        case TableDesign.MediumList1Accent4:
                            val.Value = "MediumList1-Accent4";
                            break;

                        case TableDesign.MediumList1Accent5:
                            val.Value = "MediumList1-Accent5";
                            break;

                        case TableDesign.MediumList1Accent6:
                            val.Value = "MediumList1-Accent6";
                            break;

                        case TableDesign.MediumList2:
                            val.Value = "MediumList2";
                            break;

                        case TableDesign.MediumList2Accent1:
                            val.Value = "MediumList2-Accent1";
                            break;

                        case TableDesign.MediumList2Accent2:
                            val.Value = "MediumList2-Accent2";
                            break;

                        case TableDesign.MediumList2Accent3:
                            val.Value = "MediumList2-Accent3";
                            break;

                        case TableDesign.MediumList2Accent4:
                            val.Value = "MediumList2-Accent4";
                            break;

                        case TableDesign.MediumList2Accent5:
                            val.Value = "MediumList2-Accent5";
                            break;

                        case TableDesign.MediumList2Accent6:
                            val.Value = "MediumList2-Accent6";
                            break;

                        case TableDesign.MediumGrid1:
                            val.Value = "MediumGrid1";
                            break;

                        case TableDesign.MediumGrid1Accent1:
                            val.Value = "MediumGrid1-Accent1";
                            break;

                        case TableDesign.MediumGrid1Accent2:
                            val.Value = "MediumGrid1-Accent2";
                            break;

                        case TableDesign.MediumGrid1Accent3:
                            val.Value = "MediumGrid1-Accent3";
                            break;

                        case TableDesign.MediumGrid1Accent4:
                            val.Value = "MediumGrid1-Accent4";
                            break;

                        case TableDesign.MediumGrid1Accent5:
                            val.Value = "MediumGrid1-Accent5";
                            break;

                        case TableDesign.MediumGrid1Accent6:
                            val.Value = "MediumGrid1-Accent6";
                            break;

                        case TableDesign.MediumGrid2:
                            val.Value = "MediumGrid2";
                            break;

                        case TableDesign.MediumGrid2Accent1:
                            val.Value = "MediumGrid2-Accent1";
                            break;

                        case TableDesign.MediumGrid2Accent2:
                            val.Value = "MediumGrid2-Accent2";
                            break;

                        case TableDesign.MediumGrid2Accent3:
                            val.Value = "MediumGrid2-Accent3";
                            break;

                        case TableDesign.MediumGrid2Accent4:
                            val.Value = "MediumGrid2-Accent4";
                            break;

                        case TableDesign.MediumGrid2Accent5:
                            val.Value = "MediumGrid2-Accent5";
                            break;

                        case TableDesign.MediumGrid2Accent6:
                            val.Value = "MediumGrid2-Accent6";
                            break;

                        case TableDesign.MediumGrid3:
                            val.Value = "MediumGrid3";
                            break;

                        case TableDesign.MediumGrid3Accent1:
                            val.Value = "MediumGrid3-Accent1";
                            break;

                        case TableDesign.MediumGrid3Accent2:
                            val.Value = "MediumGrid3-Accent2";
                            break;

                        case TableDesign.MediumGrid3Accent3:
                            val.Value = "MediumGrid3-Accent3";
                            break;

                        case TableDesign.MediumGrid3Accent4:
                            val.Value = "MediumGrid3-Accent4";
                            break;

                        case TableDesign.MediumGrid3Accent5:
                            val.Value = "MediumGrid3-Accent5";
                            break;

                        case TableDesign.MediumGrid3Accent6:
                            val.Value = "MediumGrid3-Accent6";
                            break;

                        case TableDesign.DarkList:
                            val.Value = "DarkList";
                            break;

                        case TableDesign.DarkListAccent1:
                            val.Value = "DarkList-Accent1";
                            break;

                        case TableDesign.DarkListAccent2:
                            val.Value = "DarkList-Accent2";
                            break;

                        case TableDesign.DarkListAccent3:
                            val.Value = "DarkList-Accent3";
                            break;

                        case TableDesign.DarkListAccent4:
                            val.Value = "DarkList-Accent4";
                            break;

                        case TableDesign.DarkListAccent5:
                            val.Value = "DarkList-Accent5";
                            break;

                        case TableDesign.DarkListAccent6:
                            val.Value = "DarkList-Accent6";
                            break;

                        case TableDesign.ColorfulShading:
                            val.Value = "ColorfulShading";
                            break;

                        case TableDesign.ColorfulShadingAccent1:
                            val.Value = "ColorfulShading-Accent1";
                            break;

                        case TableDesign.ColorfulShadingAccent2:
                            val.Value = "ColorfulShading-Accent2";
                            break;

                        case TableDesign.ColorfulShadingAccent3:
                            val.Value = "ColorfulShading-Accent3";
                            break;

                        case TableDesign.ColorfulShadingAccent4:
                            val.Value = "ColorfulShading-Accent4";
                            break;

                        case TableDesign.ColorfulShadingAccent5:
                            val.Value = "ColorfulShading-Accent5";
                            break;

                        case TableDesign.ColorfulShadingAccent6:
                            val.Value = "ColorfulShading-Accent6";
                            break;

                        case TableDesign.ColorfulList:
                            val.Value = "ColorfulList";
                            break;

                        case TableDesign.ColorfulListAccent1:
                            val.Value = "ColorfulList-Accent1";
                            break;

                        case TableDesign.ColorfulListAccent2:
                            val.Value = "ColorfulList-Accent2";
                            break;

                        case TableDesign.ColorfulListAccent3:
                            val.Value = "ColorfulList-Accent3";
                            break;

                        case TableDesign.ColorfulListAccent4:
                            val.Value = "ColorfulList-Accent4";
                            break;

                        case TableDesign.ColorfulListAccent5:
                            val.Value = "ColorfulList-Accent5";
                            break;

                        case TableDesign.ColorfulListAccent6:
                            val.Value = "ColorfulList-Accent6";
                            break;

                        case TableDesign.ColorfulGrid:
                            val.Value = "ColorfulGrid";
                            break;

                        case TableDesign.ColorfulGridAccent1:
                            val.Value = "ColorfulGrid-Accent1";
                            break;

                        case TableDesign.ColorfulGridAccent2:
                            val.Value = "ColorfulGrid-Accent2";
                            break;

                        case TableDesign.ColorfulGridAccent3:
                            val.Value = "ColorfulGrid-Accent3";
                            break;

                        case TableDesign.ColorfulGridAccent4:
                            val.Value = "ColorfulGrid-Accent4";
                            break;

                        case TableDesign.ColorfulGridAccent5:
                            val.Value = "ColorfulGrid-Accent5";
                            break;

                        case TableDesign.ColorfulGridAccent6:
                            val.Value = "ColorfulGrid-Accent6";
                            break;

                        default:
                            break;
                    }
                }
                if (Document.styles == null)
                {
                    PackagePart word_styles = Document.package.GetPart(new Uri("/word/styles.xml", UriKind.Relative));
                    using (TextReader tr = new StreamReader(word_styles.GetStream()))
                        Document.styles = XDocument.Load(tr);
                }

                var tableStyle =
                (
                    from e in Document.styles.Descendants()
                    let styleId = e.Attribute(DocxNamespace.Main + "styleId")
                    where (styleId != null && styleId.Value == val.Value)
                    select e
                ).FirstOrDefault();

                if (tableStyle == null)
                {
                    XDocument external_style_doc = Resources.DefaultTableStyles;

                    var styleElement =
                    (
                        from e in external_style_doc.Descendants()
                        let styleId = e.Attribute(DocxNamespace.Main + "styleId")
                        where (styleId != null && styleId.Value == val.Value)
                        select e
                    ).First();

                    Document.styles.Element(DocxNamespace.Main + "styles").Add(styleElement);
                }
            }
        }

        /// <summary>
        /// Get all of the Hyperlinks in this Table.
        /// </summary>
        public List<Hyperlink> Hyperlinks => Rows.SelectMany(r => r.Hyperlinks).ToList();

        /// <summary>
        /// Returns the index of this Table.
        /// </summary>
        /// <example>
        /// Replace the first table in this document with a new Table.
        /// <code>
        /// // Load a document into memory.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table t = document.Tables[0];
        ///
        ///     // Get the character index of Table t in this document.
        ///     int index = t.Index;
        ///
        ///     // Remove Table t.
        ///     t.Remove();
        ///
        ///     // Insert a new Table at the original index of Table t.
        ///     Table newTable = document.InsertTable(index, 4, 4);
        ///
        ///     // Set the design of this new Table, so that we can see it.
        ///     newTable.Design = TableDesign.LightShadingAccent1;
        ///
        ///     // Save all changes made to the document.
        ///     document.Save();
        /// } // Release this document from memory.
        /// </code>
        /// </example>
        public int Index
        {
            get
            {
                int index = 0;
                IEnumerable<XElement> previous = Xml.ElementsBeforeSelf();

                foreach (XElement e in previous)
                    index += Paragraph.GetElementTextLength(e);

                return index;
            }
        }

        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        ///
        public virtual List<Paragraph> Paragraphs => Rows.SelectMany(r => r.Paragraphs).ToList();

        /// <summary>
        /// Returns a list of all Pictures in a Table.
        /// </summary>
        public List<Picture> Pictures => Rows.SelectMany(r => r.Pictures).ToList();

        /// <summary>
        /// Returns the number of rows in this table.
        /// </summary>
        public int RowCount => Xml.Elements(DocxNamespace.Main + "tr").Count();

        /// <summary>
        /// Returns a list of rows in this table.
        /// </summary>
        public List<Row> Rows => Xml.Elements(DocxNamespace.Main + "tr").Select(r => new Row(this, Document, r)).ToList();

        /// <summary>
        /// Gets or Sets the value of the Table Caption (Alternate Text Title) of this table.
        /// </summary>
        public string TableCaption
        {
            set
            {
                XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
                if (tblPr != null)
                {
                    XElement tblCaption =
                        tblPr.Descendants(DocxNamespace.Main + "tblCaption").FirstOrDefault();

                    if (tblCaption != null)
                        tblCaption.Remove();

                    tblCaption = new XElement(DocxNamespace.Main + "tblCaption",
                        new XAttribute(DocxNamespace.Main + "val", value));
                    tblPr.Add(tblCaption);
                }
            }

            get
            {
                XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
                if (tblPr != null)
                {
                    XElement caption = tblPr.Element(DocxNamespace.Main + "tblCaption");
                    if (caption != null)
                    {
                        tableCaption = caption.AttributeValue(DocxNamespace.Main + "val");
                    }
                }
                return tableCaption;
            }
        }

        /// <summary>
        /// Gets or Sets the value of the Table Description (Alternate Text Description) of this table.
        /// </summary>
        public string TableDescription
        {
            set
            {
                XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
                if (tblPr != null)
                {
                    XElement tblDescription =
                        tblPr.Descendants(DocxNamespace.Main + "tblDescription").FirstOrDefault();

                    if (tblDescription != null)
                        tblDescription.Remove();

                    tblDescription = new XElement(DocxNamespace.Main + "tblDescription",
                       new XAttribute(DocxNamespace.Main + "val", value));
                    tblPr.Add(tblDescription);
                }
            }

            get
            {
                XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
                if (tblPr != null)
                {
                    XElement caption = tblPr.Element(DocxNamespace.Main + "tblDescription");
                    if (caption != null)
                    {
                        tableDescription = caption.AttributeValue(DocxNamespace.Main + "val");
                    }
                }
                return tableDescription;
            }
        }

        public TableLook TableLook { get; set; }

        /// <summary>
        /// Get a border edge value for this table
        /// </summary>
        /// <param name="borderType">The table border to get</param>
        public Border GetBorder(TableBorderType borderType)
        {
            Border border = new Border();

            XElement tblPr = Xml.Element(DocxNamespace.Main + "tblPr");
            if (tblPr != null)
            {
                XElement tblBorders = tblPr.Element(DocxNamespace.Main + "tblBorders");
                if (tblBorders != null)
                {
                    string borderTypeName = borderType.GetEnumName();
                    XElement tblBorderType = tblBorders.Element(DocxNamespace.Main + borderTypeName);
                    border.GetDetails(tblBorderType);
                }
            }
            return border;
        }

        /// <summary>
        /// Gets the column width for a given column index.
        /// </summary>
        /// <param name="index"></param>
        public double GetColumnWidth(int index) => ColumnWidths == null || index > ColumnWidths.Count - 1 ? double.NaN : ColumnWidths[index];

        /// <summary>
        /// Insert a column to the right of a Table.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Insert a new column to this right of this table.
        ///     table.InsertColumn();
        ///
        ///     // Set the new columns text to "Row no."
        ///     table.Rows[0].Cells[table.ColumnCount - 1].Paragraph.InsertText("Row no.", false);
        ///
        ///     // Loop through each row in the table.
        ///     for (int i = 1; i &lt; table.Rows.Count; i++)
        ///     {
        ///         // The current row.
        ///         Row row = table.Rows[i];
        ///
        ///         // The cell in this row that belongs to the new column.
        ///         Cell cell = row.Cells[table.ColumnCount - 1];
        ///
        ///         // The first Paragraph that this cell houses.
        ///         Paragraph p = cell.Paragraphs[0];
        ///
        ///         // Insert this rows index.
        ///         p.InsertText(i.ToString(), false);
        ///     }
        ///
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void InsertColumn()
        {
            InsertColumn(ColumnCount);
        }

        /// <summary>
        /// Insert a column into a table.
        /// </summary>
        /// <param name="index">The index to insert the column at.</param>
        /// <example>
        /// Insert a column to the left of a table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Insert a new column to this left of this table.
        ///     table.InsertColumn(0);
        ///
        ///     // Set the new columns text to "Row no."
        ///     table.Rows[0].Cells[table.ColumnCount - 1].Paragraph.InsertText("Row no.", false);
        ///
        ///     // Loop through each row in the table.
        ///     for (int i = 1; i &lt; table.Rows.Count; i++)
        ///     {
        ///         // The current row.
        ///         Row row = table.Rows[i];
        ///
        ///         // The cell in this row that belongs to the new column.
        ///         Cell cell = row.Cells[table.ColumnCount - 1];
        ///
        ///         // The first Paragraph that this cell houses.
        ///         Paragraph p = cell.Paragraphs[0];
        ///
        ///         // Insert this rows index.
        ///         p.InsertText(i.ToString(), false);
        ///     }
        ///
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void InsertColumn(int index)
        {
            if (RowCount > 0)
            {
                cachedColCount = -1;
                foreach (Row r in Rows)
                {
                    // create cell
                    XElement cell = HelperFunctions.CreateTableCell();

                    // insert cell
                    var cells = r.Cells.ToList();
                    if (cells.Count == index)
                        cells[index - 1].Xml.AddAfterSelf(cell);
                    else
                        cells[index].Xml.AddBeforeSelf(cell);
                }
            }
        }

        /// <summary>
        /// Insert a page break after a Table.
        /// </summary>
        /// <example>
        /// Insert a Table and a Paragraph into a document with a page break between them.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Table.
        ///     Table t1 = document.InsertTable(2, 2);
        ///     t1.Design = TableDesign.LightShadingAccent1;
        ///
        ///     // Insert a page break after this Table.
        ///     t1.InsertPageBreakAfterSelf();
        ///
        ///     // Insert a new Paragraph.
        ///     Paragraph p1 = document.InsertParagraph("Paragraph", false);
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override void InsertPageBreakAfterSelf()
        {
            base.InsertPageBreakAfterSelf();
        }

        /// <summary>
        /// Insert a page break before a Table.
        /// </summary>
        /// <example>
        /// Insert a Table and a Paragraph into a document with a page break between them.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p1 = document.InsertParagraph("Paragraph", false);
        ///
        ///     // Insert a new Table.
        ///     Table t1 = document.InsertTable(2, 2);
        ///     t1.Design = TableDesign.LightShadingAccent1;
        ///
        ///     // Insert a page break before this Table.
        ///     t1.InsertPageBreakBeforeSelf();
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override void InsertPageBreakBeforeSelf()
        {
            base.InsertPageBreakBeforeSelf();
        }

        /// <summary>
        /// Insert a Paragraph after this Table, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into document b after this Table.
        /// <code>
        /// // Place holder for a Paragraph.
        /// Paragraph p;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first paragraph from this document.
        ///     p = documentA.Paragraphs[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Table in document b.
        ///     Table t = documentB.Tables[0];
        ///
        ///     // Insert the Paragraph from document a after this Table.
        ///     Paragraph newParagraph = t.InsertParagraphAfterSelf(p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(Paragraph p)
        {
            return base.InsertParagraphAfterSelf(p);
        }

        /// <summary>
        /// Insert a new Paragraph after this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting to apply to this insertion.</param>
        /// <returns>A new Paragraph inserted after this Table.</returns>
        /// <example>
        /// Insert a new paragraph after the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     Formatting boldFormatting = new Formatting();
        ///     boldFormatting.Bold = true;
        ///
        ///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.", false, boldFormatting);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
        {
            return base.InsertParagraphAfterSelf(text, trackChanges, formatting);
        }

        /// <summary>
        /// Insert a new Paragraph after this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph inserted after this Table.</returns>
        /// <example>
        /// Insert a new paragraph after the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.", false);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
        {
            return base.InsertParagraphAfterSelf(text, trackChanges);
        }

        /// <summary>
        /// Insert a new Paragraph after this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <returns>A new Paragraph inserted after this Table.</returns>
        /// <example>
        /// Insert a new Paragraph after the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.");
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(string text)
        {
            return base.InsertParagraphAfterSelf(text);
        }

        /// <summary>
        /// Insert a Paragraph before this Table, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into document b before this Table.
        /// <code>
        /// // Place holder for a Paragraph.
        /// Paragraph p;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first paragraph from this document.
        ///     p = documentA.Paragraphs[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Table in document b.
        ///     Table t = documentB.Tables[0];
        ///
        ///     // Insert the Paragraph from document a before this Table.
        ///     Paragraph newParagraph = t.InsertParagraphBeforeSelf(p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(Paragraph p)
        {
            return base.InsertParagraphBeforeSelf(p);
        }

        /// <summary>
        /// Insert a new Paragraph before this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <returns>A new Paragraph inserted before this Table.</returns>
        /// <example>
        /// Insert a new Paragraph before the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.");
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(string text)
        {
            return base.InsertParagraphBeforeSelf(text);
        }

        /// <summary>
        /// Insert a new Paragraph before this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph inserted before this Table.</returns>
        /// <example>
        /// Insert a new paragraph before the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.", false);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
        {
            return base.InsertParagraphBeforeSelf(text, trackChanges);
        }

        /// <summary>
        /// Insert a new Paragraph before this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting to apply to this insertion.</param>
        /// <returns>A new Paragraph inserted before this Table.</returns>
        /// <example>
        /// Insert a new paragraph before the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     Formatting boldFormatting = new Formatting();
        ///     boldFormatting.Bold = true;
        ///
        ///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.", false, boldFormatting);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
        {
            return base.InsertParagraphBeforeSelf(text, trackChanges, formatting);
        }

        /// <summary>
        /// Insert a row at the end of this table.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Insert a new row at the end of this table.
        ///     Row row = table.InsertRow();
        ///
        ///     // Loop through each cell in this new row.
        ///     foreach (Cell c in row.Cells)
        ///     {
        ///         // Set the text of each new cell to "Hello".
        ///         c.Paragraphs[0].InsertText("Hello", false);
        ///     }
        ///
        ///     // Save the document to a new file.
        ///     document.SaveAs(@"C:\Example\Test2.docx");
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <returns>A new row.</returns>
        public Row InsertRow()
        {
            return InsertRow(RowCount);
        }

        /// <summary>
        /// Insert a copy of a row at the end of this table.
        /// </summary>
        /// <returns>A new row.</returns>
        public Row InsertRow(Row row)
        {
            return InsertRow(row, RowCount);
        }

        /// <summary>
        /// Insert a row into this table.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Insert a new row at index 1 in this table.
        ///     Row row = table.InsertRow(1);
        ///
        ///     // Loop through each cell in this new row.
        ///     foreach (Cell c in row.Cells)
        ///     {
        ///         // Set the text of each new cell to "Hello".
        ///         c.Paragraphs[0].InsertText("Hello", false);
        ///     }
        ///
        ///     // Save the document to a new file.
        ///     document.SaveAs(@"C:\Example\Test2.docx");
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <param name="index">Index to insert row at.</param>
        /// <returns>A new Row</returns>
        public Row InsertRow(int index)
        {
            if (index < 0 || index > RowCount)
                throw new IndexOutOfRangeException();

            List<XElement> content = new List<XElement>();
            for (int i = 0; i < ColumnCount; i++)
            {
                var w = 2310d;
                if (ColumnWidthsValue != null && ColumnWidthsValue.Length > i)
                    w = ColumnWidthsValue[i] * 15;
                XElement cell = HelperFunctions.CreateTableCell(w);
                content.Add(cell);
            }

            return InsertRow(content, index);
        }

        /// <summary>
        /// Insert a copy of a row into this table.
        /// </summary>
        /// <param name="row">Row to copy and insert.</param>
        /// <param name="index">Index to insert row at.</param>
        /// <returns>A new Row</returns>
        public Row InsertRow(Row row, int index)
        {
            if (row == null)
                throw new ArgumentNullException("row");

            if (index < 0 || index > RowCount)
                throw new IndexOutOfRangeException();

            List<XElement> content = row.Xml.Elements(DocxNamespace.Main + "tc").Select(element => HelperFunctions.CloneElement(element)).ToList();
            return InsertRow(content, index);
        }

        /// <summary>
        /// Insert a new Table after this Table, this Table can be from this document or another document.
        /// </summary>
        /// <param name="t">The Table t to be inserted</param>
        /// <returns>A new Table inserted after this Table.</returns>
        /// <example>
        /// Insert a new Table after this Table.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Table in document b.
        ///     Table t2 = documentB.Tables[0];
        ///
        ///     // Insert the Table from document a after this Table.
        ///     Table newTable = t2.InsertTableAfterSelf(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableAfterSelf(Table t)
        {
            return base.InsertTableAfterSelf(t);
        }

        /// <summary>
        /// Insert a new Table into this document after this Table.
        /// </summary>
        /// <param name="rowCount">The number of rows this Table should have.</param>
        /// <param name="columnCount">The number of columns this Table should have.</param>
        /// <returns>A new Table inserted before this Table.</returns>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     //Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///     t.Design = TableDesign.LightShadingAccent1;
        ///     t.Alignment = Alignment.center;
        ///
        ///     // Insert a new Table after this Table.
        ///     Table newTable = t.InsertTableAfterSelf(2, 2);
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///     newTable.Alignment = Alignment.center;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableAfterSelf(int rowCount, int columnCount)
        {
            return base.InsertTableAfterSelf(rowCount, columnCount);
        }

        /// <summary>
        /// Insert a new Table before this Table, this Table can be from this document or another document.
        /// </summary>
        /// <param name="t">The Table t to be inserted</param>
        /// <returns>A new Table inserted before this Table.</returns>
        /// <example>
        /// Insert a new Table before this Table.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Table in document b.
        ///     Table t2 = documentB.Tables[0];
        ///
        ///     // Insert the Table from document a before this Table.
        ///     Table newTable = t2.InsertTableBeforeSelf(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableBeforeSelf(Table t)
        {
            return base.InsertTableBeforeSelf(t);
        }

        /// <summary>
        /// Insert a new Table into this document before this Table.
        /// </summary>
        /// <param name="rowCount">The number of rows this Table should have.</param>
        /// <param name="columnCount">The number of columns this Table should have.</param>
        /// <returns>A new Table inserted before this Table.</returns>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     //Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///     t.Design = TableDesign.LightShadingAccent1;
        ///     t.Alignment = Alignment.center;
        ///
        ///     // Insert a new Table before this Table.
        ///     Table newTable = t.InsertTableBeforeSelf(2, 2);
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///     newTable.Alignment = Alignment.center;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableBeforeSelf(int rowCount, int columnCount)
        {
            return base.InsertTableBeforeSelf(rowCount, columnCount);
        }

        /// <summary>
        /// Merge cells in given column starting with startRow and ending with endRow.
        /// </summary>
        public void MergeCellsInColumn(int columnIndex, int startRow, int endRow)
        {
            // Check for valid start and end indexes.
            if (columnIndex < 0 || columnIndex >= ColumnCount)
                throw new IndexOutOfRangeException(nameof(columnIndex));

            if (startRow < 0 || endRow <= startRow || endRow >= Rows.Count)
                throw new IndexOutOfRangeException();
            
            foreach (Row row in Rows.Where((_, i) => i > startRow && i <= endRow))
            {
                Cell c = row.Cells.ElementAt(columnIndex);
                var tcPr = c.Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                _ = tcPr.GetOrCreateElement(DocxNamespace.Main + "vMerge");
            }

            var startRowCells = Rows[startRow].Cells.ToList();

            XElement start_tcPr = null;
            if (columnIndex > startRowCells.Count)
                start_tcPr = startRowCells[startRowCells.Count - 1].Xml.Element(DocxNamespace.Main + "tcPr");
            else
                start_tcPr = startRowCells[columnIndex].Xml.Element(DocxNamespace.Main + "tcPr");
            if (start_tcPr == null)
            {
                startRowCells[columnIndex].Xml.SetElementValue(DocxNamespace.Main + "tcPr", string.Empty);
                start_tcPr = startRowCells[columnIndex].Xml.Element(DocxNamespace.Main + "tcPr");
            }

            /*
			  * Get the gridSpan element of this row,
			  * null will be returned if no such element exists.
			  */
            XElement start_vMerge = start_tcPr.Element(DocxNamespace.Main + "vMerge");
            if (start_vMerge == null)
            {
                start_tcPr.SetElementValue(DocxNamespace.Main + "vMerge", string.Empty);
                start_vMerge = start_tcPr.Element(DocxNamespace.Main + "vMerge");
            }

            start_vMerge.SetAttributeValue(DocxNamespace.Main + "val", "restart");
        }
        /// <summary>
        /// Remove this Table from this document.
        /// </summary>
        /// <example>
        /// Remove the first Table from this document.
        /// <code>
        /// // Load a document into memory.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table t = d.Tables[0];
        ///
        ///     // Remove this Table.
        ///     t.Remove();
        ///
        ///     // Save all changes made to the document.
        ///     document.Save();
        /// } // Release this document from memory.
        /// </code>
        /// </example>
        public void Remove()
        {
            Xml.Remove();
        }

        /// <summary>
        /// Remove the last column for this Table.
        /// </summary>
        /// <example>
        /// Remove the last column from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the last column from this table.
        ///     table.RemoveColumn();
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveColumn()
        {
            RemoveColumn(ColumnCount - 1);
        }

        /// <summary>
        /// Remove a column from this Table.
        /// </summary>
        /// <param name="index">The column to remove.</param>
        /// <example>
        /// Remove the first column from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the first column from this table.
        ///     table.RemoveColumn(0);
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveColumn(int index)
        {
            if (index < 0 || index > ColumnCount - 1)
                throw new IndexOutOfRangeException();

            foreach (Row r in Rows)
            {
                r.Cells.ElementAt(index).Xml.Remove();
            }
            
            cachedColCount = -1;
        }

        /// <summary>
        /// Remove the last row from this Table.
        /// </summary>
        /// <example>
        /// Remove the last row from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the last row from this table.
        ///     table.RemoveRow();
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveRow()
        {
            RemoveRow(RowCount - 1);
        }

        /// <summary>
        /// Remove a row from this Table.
        /// </summary>
        /// <param name="index">The row to remove.</param>
        /// <example>
        /// Remove the first row from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the first row from this table.
        ///     table.RemoveRow(0);
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveRow(int index)
        {
            if (index < 0 || index > RowCount - 1)
                throw new IndexOutOfRangeException();

            Rows[index].Xml.Remove();
            if (Rows.Count == 0)
                Remove();
        }

        /// <summary>
        /// Set a table border
        /// Added by lckuiper @ 20101117
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a new document.
        ///using (DocX document = DocX.Create("Test.docx"))
        ///{
        ///    // Insert a table into this document.
        ///    Table t = document.InsertTable(3, 3);
        ///
        ///    // Create a large blue border.
        ///    Border b = new Border(BorderStyle.Tcbs_single, BorderSize.seven, 0, Color.Blue);
        ///
        ///    // Set the tables Top, Bottom, Left and Right Borders to b.
        ///    t.SetBorder(TableBorderType.Top, b);
        ///    t.SetBorder(TableBorderType.Bottom, b);
        ///    t.SetBorder(TableBorderType.Left, b);
        ///    t.SetBorder(TableBorderType.Right, b);
        ///
        ///    // Save the document.
        ///    document.Save();
        ///}
        /// </code>
        /// </example>
        /// <param name="borderType">The table border to set</param>
        /// <param name="border">Border object to set the table border</param>
        public void SetBorder(TableBorderType borderType, Border border)
        {
            var tblPrName = DocxNamespace.Main + "tblPr";
            var tblBordersName = DocxNamespace.Main + "tblBorders";

            // Get the table properties element for this Table
            XElement tblPr = Xml.Element(tblPrName);
            if (tblPr == null)
            {
                // Create
                Xml.SetElementValue(tblPrName, string.Empty);
                tblPr = Xml.Element(tblPrName);
            }

            // Get the table borders element for this Table,
            XElement tblBorders = tblPr.Element(tblBordersName);
            if (tblBorders == null)
            {
                tblPr.SetElementValue(tblBordersName, string.Empty);
                tblBorders = tblPr.Element(tblBordersName);
            }

            // Get the 'borderType' element for this Table
            string btValue = borderType.ToString();
            if (!string.IsNullOrEmpty(btValue) && char.IsUpper(btValue[0]))
                btValue = char.ToLower(btValue[0]) + btValue.Substring(1);

            var btValueName = DocxNamespace.Main + btValue;
            XElement tblBorderType = tblBorders.Element(btValueName);
            if (tblBorderType == null)
            {
                tblBorders.SetElementValue(btValueName, string.Empty);
                tblBorderType = tblBorders.Element(btValueName);
            }

            // get string value of border style
            string borderstyle = border.Style.GetEnumName();

            // The val attribute is used for the border style
            tblBorderType.SetAttributeValue(DocxNamespace.Main + "val", borderstyle);

            if (border.Style != BorderStyle.Empty)
            {
                var size = border.Size switch
                {
                    BorderSize.One => 2,
                    BorderSize.Two => 4,
                    BorderSize.Three => 6,
                    BorderSize.Four => 8,
                    BorderSize.Five => 12,
                    BorderSize.Six => 18,
                    BorderSize.Seven => 24,
                    BorderSize.Eight => 36,
                    BorderSize.Nine => 48,
                    _ => 2,
                };

                // The sz attribute is used for the border size
                tblBorderType.SetAttributeValue(DocxNamespace.Main + "sz", size.ToString());

                // The space attribute is used for the cell spacing (probably '0')
                tblBorderType.SetAttributeValue(DocxNamespace.Main + "space", border.Space.ToString());

                // The color attribute is used for the border color
                tblBorderType.SetAttributeValue(DocxNamespace.Main + "color", border.Color.ToHex());
            }
        }

        /// <summary>
        /// Sets the column width for the given index.
        /// </summary>
        /// <param name="index">Column index</param>
        /// <param name="width">Colum width</param>
        public void SetColumnWidth(int index, double width)
        {
            var columnWidths = ColumnWidths;

            if (columnWidths == null || index > columnWidths.Count - 1)
            {
                if (Rows.Count == 0)
                    throw new Exception("No rows available.");

                columnWidths = new List<double>();
                foreach (Cell c in Rows[Rows.Count - 1].Cells)
                {
                    columnWidths.Add(c.Width);
                }
            }

            // check if index is matching table columns
            if (index > columnWidths.Count - 1)
                throw new Exception("The index is greather than the available table columns.");

            // get the table grid props
            XElement grid = Xml.Element(DocxNamespace.Main + "tblGrid");
            if (grid == null)
            {
                XElement tblPr = GetOrCreateTablePropertiesSection();
                tblPr.AddAfterSelf(new XElement(DocxNamespace.Main + "tblGrid"));
                grid = Xml.Element(DocxNamespace.Main + "tblGrid");
            }

            // remove all existing values
            grid.RemoveAll();

            for (int i = 0; i < columnWidths.Count; i++)
            {
                double value = (i == index) ? width : columnWidths[i];
                grid.Add(new XElement(DocxNamespace.Main + "gridCol",
                            new XAttribute(DocxNamespace.Main + "w", value)));
            }

            // remove cell widths
            foreach (Row r in Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    c.Width = -1;
                }
            }

            // set fitting to fixed; this will add/set additional table properties
            this.AutoFit = AutoFit.Fixed;
        }

        /// <summary>
        /// Set the direction of all content in this Table.
        /// </summary>
        /// <param name="direction">(Left to Right) or (Right to Left)</param>
        public void SetDirection(Direction direction)
        {
            XElement tblPr = GetOrCreateTablePropertiesSection();
            tblPr.Add(new XElement(DocxNamespace.Main + "bidiVisual"));
            Rows.ForEach(r => r.SetDirection(direction));
        }
        /// <summary>
        /// Set the specified cell margin for the table-level.
        /// </summary>
        /// <param name="type">The side of the cell margin.</param>
        /// <param name="margin">The value for the specified cell margin.</param>
        /// <remarks>More information can be found <see cref="http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellmargindefault.aspx">here</see></remarks>
        public void SetTableCellMargin(TableCellMarginType type, double margin)
        {
            XElement tblPr = GetOrCreateTablePropertiesSection();

            // find (or create) the element with the cell margins
            XElement tblCellMar = tblPr.Element(DocxNamespace.Main + "tblCellMar");
            if (tblCellMar == null)
            {
                tblPr.AddFirst(new XElement(DocxNamespace.Main + "tblCellMar"));
                tblCellMar = tblPr.Element(DocxNamespace.Main + "tblCellMar");
            }

            // find (or create) the element with cell margin for the specified side
            XElement tblMargin = tblCellMar.Element(DocxNamespace.Main + type.ToString());
            if (tblMargin == null)
            {
                tblCellMar.AddFirst(new XElement(DocxNamespace.Main + type.ToString()));
                tblMargin = tblCellMar.Element(DocxNamespace.Main + type.ToString());
            }

            tblMargin.RemoveAttributes();
            // set the value for the cell margin
            tblMargin.Add(new XAttribute(DocxNamespace.Main + "w", margin));
            // set the side of cell margin
            tblMargin.Add(new XAttribute(DocxNamespace.Main + "type", "dxa"));
        }

        /// <summary>
        /// Set the widths for the columns
        /// </summary>
        /// <param name="widths"></param>
        public void SetWidths(float[] widths)
        {
            ColumnWidthsValue = widths;
            
            foreach (var row in Rows)
            {
                for (var col = 0; col < widths.Length; col++)
                {
                    var rowCells = row.Cells.ToList();
                    if (rowCells.Count > col)
                        rowCells[col].Width = widths[col];
                }
            }
        }

        /// <summary>
        /// Retrieves or create the table properties (tblPr) section in the document.
        /// </summary>
        /// <returns>The tblPr element for this Table.</returns>
        internal XElement GetOrCreateTablePropertiesSection()
        {
            var tblPrName = DocxNamespace.Main + "tblPr";

            XElement tblPr = Xml.Element(tblPrName);
            if (tblPr == null)
            {
                Xml.AddFirst(new XElement(tblPrName));
                tblPr = Xml.Element(tblPrName);
            }

            return tblPr;
        }
        private Row InsertRow(List<XElement> content, Int32 index)
        {
            Row newRow = new Row(this, Document, new XElement(DocxNamespace.Main + "tr", content));

            XElement rowXml;
            if (index == Rows.Count)
            {
                rowXml = Rows.Last().Xml;
                rowXml.AddAfterSelf(newRow.Xml);
            }
            else
            {
                rowXml = Rows[index].Xml;
                rowXml.AddBeforeSelf(newRow.Xml);
            }

            return newRow;
        }
    }
    public class TableLook
    {
        public bool FirstColumn { get; set; }
        public bool FirstRow { get; set; }
        public bool LastColumn { get; set; }
        public bool LastRow { get; set; }
        public bool NoHorizontalBanding { get; set; }
        public bool NoVerticalBanding { get; set; }
    }
}