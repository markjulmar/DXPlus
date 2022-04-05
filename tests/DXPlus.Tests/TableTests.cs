using System;
using System.Linq;
using System.Xml.XPath;
using DXPlus.Internal;
using Xunit;

namespace DXPlus.Tests
{
    public class TableTests
    {
        [Fact]
        public void CheckExistingTablePropertyAfterInsert()
        {
            using var doc = Document.Create();

            var p = doc.Add("This is a test paragraph");
            Assert.Null(p.Table);
            doc.Add(new Table(1, 1));
            Assert.NotNull(p.Table);
        }
        
        [Fact]
        public void SeparatorParagraphInsertedBetweenTables()
        {
            using var doc = Document.Create();

            Assert.Empty(doc.Paragraphs);
            doc.Add(new Table(1, 1));
            Assert.Single(doc.Paragraphs);

            doc.Add(new Table(2, 2));
            Assert.Single(doc.Blocks.OfType<Paragraph>());

            Assert.IsType<Table>(doc.Blocks.First());
            Assert.IsType<Paragraph>(doc.Blocks.ElementAt(1));
            Assert.IsType<Table>(doc.Blocks.Last());
        }

        [Fact]
        public void TableRespectsExplicitWidthRequest()
        {
            Table t = new Table(1, 4) {Properties = new() { TableWidth = Uom.FromInches(4)} };

            double expectedSize = 4 * 1440.0;
            Assert.Equal(expectedSize, t.Properties.TableWidth.Width);
            Assert.True(double.IsNaN(t.DefaultColumnWidths.First()));

            var doc = Document.Create();
            doc.Add(t);

            Assert.False(double.IsNaN(t.DefaultColumnWidths.First()));
            Assert.Equal(expectedSize / 4, t.DefaultColumnWidths.First());
            Assert.Equal(expectedSize / 4, t.Rows.First().Cells[0].Properties.CellWidth?.Width);
            Assert.Equal(TableWidthUnit.Dxa, t.Rows.First().Cells[0].Properties.CellWidth?.Type);
        }

        [Fact]
        public void TableWithValuesAddsColRows()
        {
            var t = new Table(new[,]
            {
                { "Heading 1", "Heading 2", "Heading 3", "Heading 4" },
                { "One", "Two", "Three", "Four" },
                { "One", "Two", "Three", "Four" },
                { "One", "Two", "Three", "Four" },
                { "One", "Two", "Three", "Four" }
            });

            Assert.Equal(4, t.ColumnCount);
            Assert.Equal(5, t.Rows.Count);
            Assert.Equal("Heading 3", t.Rows[0].Cells[2].Text);
            Assert.Equal("Four", t.Rows[4].Cells[3].Text);
        }

        [Fact]
        public void TableWidthConvertersUseDxa()
        {
            var t = new Table
            {
                Properties = new() { TableWidth = 100 }
            };

            Assert.Equal(TableWidthUnit.Dxa, t.Properties.TableWidth!.Type);
            Assert.Equal(100, t.Properties.TableWidth.Width);
        }

        [Fact]
        public void CanSetPercentageWidthOnTable()
        {
            var t = new Table(1, 1)
            {
                Properties = new() {TableWidth = TableElementWidth.FromPercent(50)}
            };

            Assert.Equal(2500, t.Properties.TableWidth.Width);

            t.Properties.TableWidth = TableElementWidth.FromPercent(100);
            Assert.Equal(5000, t.Properties.TableWidth.Width);
        }

        [Fact]
        public void CreateTableWithInvalidRowsOrColumnsThrowsException()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new Table(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => new Table(1, 0));
            _ = new Table(1,1);
        }

        [Fact]
        public void TableLayoutDefaultsToNull()
        {
            var table = new Table();
            Assert.Null(table.Properties.TableLayout);
        }

        [Fact]
        public void CannotAddEmptyTableToDoc()
        {
            var doc = Document.Create();

            var table = new Table();
            Assert.Throws<Exception>(() => doc.Add(table));
        }

        [Fact]
        public void TableWithSingleRowCanBeAddedToDoc()
        {
            var doc = Document.Create();

            var table = new Table(1,1);
            _ = doc.Add(table);
        }

        [Fact]
        public void AlignmentSetsElementValue()
        {
            Table t = new Table(1,1);

            Assert.Null(t.Properties.Alignment);
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc"));

            t.Properties.Alignment = Alignment.Center;
            Assert.Equal(Alignment.Center, t.Properties.Alignment);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc"));

            t.Properties.Alignment = Alignment.Right;
            Assert.Equal(Alignment.Right, t.Properties.Alignment);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc[@val='right']"));
        }

        [Fact]
        public void AutoFitSetsElementValue()
        {
            Table t = new Table(1, 1);

            Assert.Null(t.Properties.TableLayout);

            t.AutoFit();
            Assert.Equal(TableLayout.AutoFit, t.Properties.TableLayout);

            var e = t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblLayout[@type='autofit']");
            Assert.Single(e);
        }

        [Fact]
        public void ChangingColumnWidthsOnTableAffectsCells()
        {
            Table t = new Table(1, 1);

            t.SetColumnWidths(new[] { 120.0 });
            Assert.Null(t.Properties.TableLayout);

            Assert.Equal(120.0, t.Rows.First().Cells[0].Properties.CellWidth?.Width);
            Assert.Equal(120.0, t.DefaultColumnWidths.First());
        }

        [Fact]
        public void TableDesignAddsAndRemovesElements()
        {
            Table t = new(1,1);

            Assert.Equal(TableDesign.Grid, t.Properties.Design);

            t.Properties.Design = TableDesign.ColorfulGrid;
            Assert.Equal(TableDesign.ColorfulGrid, t.Properties.Design);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle[@val='ColorfulGrid']"));

            t.Properties.Design = null;
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle"));

            t.Properties.Design = "TestDesign";
            Assert.Equal("TestDesign", t.Properties.Design);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle[@val='TestDesign']"));
        }

        [Fact]
        public void OrphanTableHasNoColumnSizes()
        {
            Table t = new Table(1, 4);

            Assert.Equal(4, t.DefaultColumnWidths.Count());
            double width = t.DefaultColumnWidths.First();
            Assert.True(double.IsNaN(width));
        }

        [Fact]
        public void AddToDocumentSetsColumnWidths()
        {
            Table t = new Table(1, 4);
            var doc = Document.Create();

            doc.Add(t);
            Assert.Equal(4, t.DefaultColumnWidths.Count());
            double width = t.DefaultColumnWidths.First();
            Assert.False(double.IsNaN(width));
        }

        [Fact]
        public void SetTableCellWidthSetsCorrectElement()
        {
            var t = new Table(1, 1);

            var width = TableElementWidth.FromPercent(100);
            Assert.Equal(5000, width.Width);
            Assert.Equal("tblW", width.Xml.Name.LocalName);

            var cell = t.Rows.First().Cells[0];
            cell.Properties.CellWidth = width;
            Assert.NotEqual(width, cell.Properties.CellWidth);
            Assert.Equal(5000, cell.Properties.CellWidth.Width);
            Assert.Single(cell.Xml.RemoveNamespaces().XPathSelectElements("//tcW"));
            Assert.Empty(cell.Xml.RemoveNamespaces().XPathSelectElements("//tblW"));
        }

        [Fact]
        public void AddToDocumentSetsUniformColumnWidths()
        {
            Table t = new Table(1, 4);
            var doc = Document.Create();

            doc.Add(t);
            Assert.Equal(4, t.DefaultColumnWidths.Count());
            double width = t.DefaultColumnWidths.First();
            var rows = t.Rows.ToList();

            Assert.Equal(width, rows[0].Cells[0].Properties.CellWidth?.Width);
            Assert.Equal(TableWidthUnit.Dxa, rows[0].Cells[0].Properties.CellWidth?.Type);
            Assert.Equal(width, rows[0].Cells[1].Properties.CellWidth?.Width);
            Assert.Equal(TableWidthUnit.Dxa, rows[0].Cells[1].Properties.CellWidth?.Type);
            Assert.Equal(width, rows[0].Cells[2].Properties.CellWidth?.Width);
            Assert.Equal(TableWidthUnit.Dxa, rows[0].Cells[2].Properties.CellWidth?.Type);
            Assert.Equal(width, rows[0].Cells[3].Properties.CellWidth?.Width);
            Assert.Equal(TableWidthUnit.Dxa, rows[0].Cells[3].Properties.CellWidth?.Type);

            Assert.Equal(width, t.DefaultColumnWidths.First());
            Assert.Equal(width, t.DefaultColumnWidths.ElementAt(1));
            Assert.Equal(width, t.DefaultColumnWidths.ElementAt(2));
            Assert.Equal(width, t.DefaultColumnWidths.Last());
        }

        [Fact]
        public void AddToDocFillsInColumnWidth()
        {
            var doc = Document.Create();
            var t = new Table(2, 2);
            Assert.True(t.DefaultColumnWidths.All(double.IsNaN));
            doc.Add(t);
            Assert.DoesNotContain(t.DefaultColumnWidths, double.IsNaN);
            Assert.True(t.DefaultColumnWidths.All(c => c > 0));
        }

        [Fact]
        public void NewTableCreatesUniformColumns()
        {
            var t = new Table(1, 4);

            Assert.Equal(4, t.DefaultColumnWidths.Count());

            double width = t.DefaultColumnWidths.First();
            Assert.Equal(width, t.DefaultColumnWidths.ElementAt(0));
            Assert.Equal(width, t.DefaultColumnWidths.ElementAt(1));
            Assert.Equal(width, t.DefaultColumnWidths.ElementAt(2));
            Assert.Equal(width, t.DefaultColumnWidths.ElementAt(3));

            Assert.Equal(TableWidthUnit.Auto, t.Properties.TableWidth?.Type);
            Assert.Equal(0, t.Properties.TableWidth?.Width);
        }

        [Fact]
        public void AddRowAddsToEndTable()
        {
            Table t = new Table(1,1);

            t.Rows.First().Cells[0].Text = "1";
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr"));

            t.AddRow().Cells[0].Text = "2";
            var rows = t.Xml.RemoveNamespaces().XPathSelectElements("//tr").ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("1", t.Rows.First().Cells[0].Paragraphs.First().Text);
            Assert.Equal("2", t.Rows.ElementAt(1).Cells[0].Paragraphs.First().Text);

            var lastRow = t.Rows.Last();
            Assert.Equal("2", lastRow.Cells[0].Text);
            Assert.Equal(1, lastRow.Cells.Count);

            t.AddRow().Cells[0].Text = "3";
            rows = t.Xml.RemoveNamespaces().XPathSelectElements("//tr").ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal("3", rows.Last().Value);
        }

        [Fact]
        public void InsertRowAtBeginning()
        {
            Table t = new Table(1, 1);

            t.Rows.First().Cells[0].Text = "2";
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr"));

            t.InsertRow(0).Cells[0].Text = "1";
            var rows = t.Xml.RemoveNamespaces().XPathSelectElements("//tr").ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("1", t.Rows.First().Cells[0].Paragraphs.First().Text);
            Assert.Equal("1", rows[0].Value);
        }

        [Fact]
        public void ConditionalFormattingParsesCorrectly()
        {
            Table t = new Table(1, 1);
            t.Properties.ConditionalFormatting = TableConditionalFormatting.FirstColumn | TableConditionalFormatting.FirstRow | TableConditionalFormatting.NoColumnBand;

            string val = t.Xml.RemoveNamespaces().XPathSelectElement("//tblLook").GetVal();
            Assert.Equal("04A0", val);
        }

        [Fact]
        public void ConditionalFormattingParsesCorrectlyBothDirections()
        {
            var flags = TableConditionalFormatting.FirstColumn | TableConditionalFormatting.FirstRow | TableConditionalFormatting.NoColumnBand;

            Table t = new Table(1, 1);
            t.Properties.ConditionalFormatting = flags;

            Assert.Equal(flags, t.Properties.ConditionalFormatting);
        }

        [Fact]
        public void ConditionalFormattingReturnsNoneOnBadValue()
        {
            Table t = new Table(1, 1);
            var e = t.Xml.Element(Namespace.Main + "tblPr").Element(Namespace.Main + "tblLook");
            e.SetAttributeValue(Name.MainVal, "badvalue");
            Assert.Equal(TableConditionalFormatting.None, t.Properties.ConditionalFormatting);
        }

        [Fact]
        public void TableStartsWithNoContent()
        {
            Table t = new Table(1, 1);

            Assert.Single(t.Rows.SelectMany(r => r.Cells).SelectMany(c => c.Paragraphs));

            var rows = t.Xml.RemoveNamespaces().XPathSelectElements("//tr").ToList();
            Assert.Single(rows);
            Assert.Empty(rows[0].Value);
        }

        [Fact]
        public void InsertRowInBetween()
        {
            Table t = new Table(2, 1);

            t.InsertRow(1).Cells[0].Text = "test";

            var rows = t.Xml.RemoveNamespaces().XPathSelectElements("//tr").ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal("test", t.Rows.ElementAt(1).Cells[0].Paragraphs.First().Text);
            Assert.Equal("test", rows[1].Value);
        }

        [Fact]
        public void SetInitialText()
        {
            Table t = new Table(1, 1);

            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr/tc/p"));
            Assert.Equal(string.Empty, t.Rows.First().Cells[0].Text);

            t.Rows.First().Cells[0].Text = "Hello";
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr/tc/p"));
            Assert.Equal("Hello", t.Rows.First().Cells[0].Text);
        }

        [Fact]
        public void TableMarginSetsElements()
        {
            Table t = new Table(1, 1);

            // Defaults
            Assert.Null(t.Properties.GetDefaultCellMargin(TableCellMarginType.Left));
            Assert.Null(t.Properties.GetDefaultCellMargin(TableCellMarginType.Top));
            Assert.Null(t.Properties.GetDefaultCellMargin(TableCellMarginType.Right));
            Assert.Null(t.Properties.GetDefaultCellMargin(TableCellMarginType.Bottom));

            t.Properties.SetDefaultCellMargin(TableCellMarginType.Top, 100);
            Assert.Equal(100, t.Properties.GetDefaultCellMargin(TableCellMarginType.Top));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblCellMar"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblCellMar/top[@w='100' and @type='dxa']"));

            t.Properties.SetDefaultCellMargin(TableCellMarginType.Top, null);
            Assert.Null(t.Properties.GetDefaultCellMargin(TableCellMarginType.Top));
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblCellMar"));
        }

        [Fact]
        void RemoveRowDeletesElement()
        {
            Table t = new Table(4, 1);

            Assert.Equal(4, t.Xml.RemoveNamespaces().XPathSelectElements("//tr").Count());

            var rows = t.Rows.ToList();
            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].Cells[0].Text = (i + 1).ToString();
            }

            Assert.Equal("3", rows[2].Cells[0].Text);

            t.RemoveRow(2);
            Assert.Equal(3, t.Xml.RemoveNamespaces().XPathSelectElements("//tr").Count());
            Assert.Equal("4", t.Rows.ElementAt(2).Cells[0].Text);
        }

        [Fact]
        void RemoveColumnDeletesElement()
        {
            Table t = new Table(rows: 1, columns: 4);

            Assert.Equal(4, t.Xml.RemoveNamespaces().XPathSelectElements("//tc").Count());
            Assert.Equal(4, t.ColumnCount);

            TableRow row = t.Rows.First();
            Assert.Equal(4, row.Cells.Count);

            for (var cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                row.Cells[cellIndex].Text = (cellIndex + 1).ToString();
            }

            Assert.Equal("2", row.Cells[1].Text);

            t.RemoveColumn(1);

            row = t.Rows.First();
            Assert.Equal("3", row.Cells[1].Text);
            Assert.Equal(3, row.Cells.Count);
            Assert.Equal(3, t.ColumnCount);
        }

        [Fact]
        void RemoveLastRowRemovesTable()
        {
            Table t = new Table(2,1);
            Assert.Equal(2, t.Xml.RemoveNamespaces().XPathSelectElements("//tr").Count());

            t.RemoveRow(1);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr"));

            t.RemoveRow(0);
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tr"));
            Assert.Null(t.Xml.Parent);
        }

        [Fact]
        void RemoveLastColumnRemovesTable()
        {
            Table t = new Table(1, 2);
            Assert.Equal(2, t.Xml.RemoveNamespaces().XPathSelectElements("//tc").Count());

            t.RemoveColumn(1);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tc"));

            t.RemoveColumn(0);
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tc"));
            Assert.Null(t.Xml.Parent);
        }

        [Fact]
        public void BreakAcrossPagesIsDefault()
        {
            Table t = new Table(2,2);

            TableRow row = t.Rows.First();
            Assert.True(row.Properties.BreakAcrossPages);
            Assert.Empty(row.Xml.RemoveNamespaces().XPathSelectElements("//trPr/cantSplit"));

            row.Properties.BreakAcrossPages = false;
            Assert.False(row.Properties.BreakAcrossPages);
            Assert.Single(row.Xml.RemoveNamespaces().XPathSelectElements("//trPr/cantSplit"));
        }

        [Fact]
        public void MergeCellsAffectGridSpan()
        {
            Table t = new Table(new[,] {{ "1", "2", "3", "4"}});

            // Merge cells 2-4
            t.Rows[0].MergeCells(1,3);

            Assert.Equal(3, t.Rows[0].Cells[1].Properties.GridSpan);
            Assert.Equal(4, t.ColumnCount);
            Assert.Equal(4, t.Rows[0].ColumnCount);
            Assert.Equal("2\n3\n4", t.Rows[0].Cells[1].Text);
            Assert.Equal(2, t.Rows[0].Xml.RemoveNamespaces().XPathSelectElements("//tc").Count());
            Assert.Single(t.Rows[0].Xml.RemoveNamespaces().XPathSelectElements("//tc/tcPr/gridSpan[@val='3']"));
        }

        [Fact]
        public void MergeCellsInColumnSetsValign()
        {
            Table t = new Table(new[,]
            {
                { "1", "2", "3"},
                { "4", "5", "6"},
                { "7", "8", "9"},
                { "10", "11", "12"},
            });

            t.MergeCellsInColumn(1, 0, t.Rows.Count);

            Assert.Equal(4, t.Rows.Count);
            Assert.Single(t.Rows[0].Xml.RemoveNamespaces().XPathSelectElements("//tc/tcPr/vMerge[@val='restart']"));
            Assert.Single(t.Rows[1].Xml.RemoveNamespaces().XPathSelectElements("//tc/tcPr/vMerge"));
            Assert.Single(t.Rows[2].Xml.RemoveNamespaces().XPathSelectElements("//tc/tcPr/vMerge"));
            Assert.Single(t.Rows[3].Xml.RemoveNamespaces().XPathSelectElements("//tc/tcPr/vMerge"));
            Assert.Equal("2\n5\n8\n11", t.Rows[0].Cells[1].Text);

            Assert.Equal(string.Empty, t.Rows[1].Cells[1].Text);
            Assert.Equal(string.Empty, t.Rows[2].Cells[1].Text);
            Assert.Equal(string.Empty, t.Rows[3].Cells[1].Text);
        }

        [Fact]
        public void PackagePartSetWhenAddedToDoc()
        {
            Table t = new Table(1,1);
            Assert.Throws<InvalidOperationException>(() => t.PackagePart);

            var doc = Document.Create();
            doc.Add(t);
            Assert.NotNull(t.PackagePart);

            var t2 = doc.Add(new Table(1, 1));
            Assert.NotNull(t2.PackagePart);
            Assert.Equal(t2.PackagePart, t.PackagePart);
            Assert.Equal(2, doc.Tables.Count());
        }
    }
}
