using System;
using System.Linq;
using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class TableTests
    {
        /*
         New Table property allows test to work
        [Fact]
        public void AddTableToOrphanParagraphFails()
        {
            FirstParagraph p = new FirstParagraph();
            Table t = new Table(1, 1);

            Assert.Throws<InvalidOperationException>(() => p.Append(t));
        }
        */

        [Fact]
        public void CreateTableWithInvalidRowsOrColumnsThrowsException()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new Table(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => new Table(1, 0));
            Table t= new Table(1,1);
        }

        [Fact]
        public void AlignmentSetsElementValue()
        {
            Table t = new Table(1,1);

            Assert.Equal(Alignment.Left, t.Alignment);
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc"));

            t.Alignment = Alignment.Center;
            Assert.Equal(Alignment.Center, t.Alignment);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc"));

            t.Alignment = Alignment.Right;
            Assert.Equal(Alignment.Right, t.Alignment);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/jc[@val='right']"));
        }

        [Fact]
        public void AutoFitSetsElementValue()
        {
            Table t = new Table(1, 1);

            Assert.Equal(AutoFit.Fixed, t.AutoFit);

            t.AutoFit = AutoFit.Contents;
            Assert.Equal(AutoFit.Contents, t.AutoFit);

            // Can't set fixed without setting column widths
            Assert.Throws<InvalidOperationException>(() => t.AutoFit = AutoFit.Fixed);

            t.SetColumnWidths(new [] { 100.0 });
            t.AutoFit = AutoFit.Fixed;
            Assert.Equal(AutoFit.Fixed, t.AutoFit);
        }

        [Fact]
        public void ChangingColumnWidthsOnTableAffectsCells()
        {
            Table t = new Table(1, 1);

            Assert.Equal(8192, t.Rows[0].Cells[0].Width);

            t.SetColumnWidths(new[] { 120.0 });
            Assert.Equal(AutoFit.Fixed, t.AutoFit);

            Assert.Equal(120.0, t.Rows[0].Cells[0].Width);
            Assert.Equal(120.0, t.GetColumnWidth(0));
        }

        [Fact]
        public void TableDesignAddsAndRemovesElements()
        {
            Table t = new Table(1,1);

            Assert.Equal(TableDesign.TableGrid, t.Design);

            t.Design = TableDesign.ColorfulGrid;
            Assert.Equal(TableDesign.ColorfulGrid, t.Design);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle[@val='ColorfulGrid']"));

            t.Design = TableDesign.None;
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle"));

            t.CustomTableDesignName = "TestDesign";
            Assert.Equal(TableDesign.Custom, t.Design);
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblStyle[@val='TestDesign']"));
        }

        [Fact]
        public void AddRowAddsToEndTable()
        {
            Table t = new Table(1,1);

            t.Rows[0].Cells[0].Text = "1";
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr"));

            t.AddRow().Cells[0].Text = "2";
            var rows = t.Xml.RemoveNamespaces().XPathSelectElements("//tr").ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("2", t.Rows[1].Cells[0].Paragraphs.First().Text);
            Assert.Equal("2", rows.Last().Value);
        }

        [Fact]
        public void InsertRowAtBeginning()
        {
            Table t = new Table(1, 1);

            t.Rows[0].Cells[0].Text = "2";
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr"));

            t.InsertRow(0).Cells[0].Text = "1";
            var rows = t.Xml.RemoveNamespaces().XPathSelectElements("//tr").ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("1", t.Rows[0].Cells[0].Paragraphs.First().Text);
            Assert.Equal("1", rows[0].Value);
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
            Assert.Equal("test", t.Rows[1].Cells[0].Paragraphs.First().Text);
            Assert.Equal("test", rows[1].Value);
        }

        [Fact]
        public void SetInitialText()
        {
            Table t = new Table(1, 1);

            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr/tc/p"));
            Assert.Equal(string.Empty, t.Rows[0].Cells[0].Text);

            t.Rows[0].Cells[0].Text = "Hello";
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tr/tc/p"));
            Assert.Equal("Hello", t.Rows[0].Cells[0].Text);
        }

        [Fact]
        public void TableMarginSetsElements()
        {
            Table t = new Table(1, 1);

            // Defaults
            Assert.Null(t.GetDefaultCellMargin(TableCellMarginType.Left));
            Assert.Null(t.GetDefaultCellMargin(TableCellMarginType.Top));
            Assert.Null(t.GetDefaultCellMargin(TableCellMarginType.Right));
            Assert.Null(t.GetDefaultCellMargin(TableCellMarginType.Bottom));

            t.SetDefaultCellMargin(TableCellMarginType.Top, 100);
            Assert.Equal(100, t.GetDefaultCellMargin(TableCellMarginType.Top));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblCellMar"));
            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblCellMar/top[@w='2000' and @type='dxa']"));

            t.SetDefaultCellMargin(TableCellMarginType.Top, null);
            Assert.Null(t.GetDefaultCellMargin(TableCellMarginType.Top));
            Assert.Empty(t.Xml.RemoveNamespaces().XPathSelectElements("//tblPr/tblCellMar"));
        }

        [Fact]
        void RemoveRowDeletesElement()
        {
            Table t = new Table(4, 1);

            Assert.Equal(4, t.Xml.RemoveNamespaces().XPathSelectElements("//tr").Count());

            t.Rows[0].Cells[0].Text = "1";
            t.Rows[1].Cells[0].Text = "2";
            t.Rows[2].Cells[0].Text = "3";
            t.Rows[3].Cells[0].Text = "4";
            Assert.Equal("3", t.Rows[2].Cells[0].Text);

            t.RemoveRow(2);
            Assert.Equal(3, t.Xml.RemoveNamespaces().XPathSelectElements("//tr").Count());
            Assert.Equal("4", t.Rows[2].Cells[0].Text);
        }

        [Fact]
        void RemoveColumnDeletesElement()
        {
            Table t = new Table(1, 4);

            Assert.Equal(4, t.Xml.RemoveNamespaces().XPathSelectElements("//tc").Count());

            for (var cellIndex = 0; cellIndex < t.Rows[0].Cells.Count; cellIndex++)
            {
                t.Rows[0].Cells[cellIndex].Text = (cellIndex + 1).ToString();
            }

            Assert.Equal("2", t.Rows[0].Cells[1].Text);

            t.RemoveColumn(1);

            Assert.Equal("3", t.Rows[0].Cells[1].Text);
            Assert.Equal(3, t.Rows[0].Cells.Count);
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

            Assert.True(t.Rows[0].BreakAcrossPages);
            Assert.Empty(t.Rows[0].Xml.RemoveNamespaces().XPathSelectElements("//trPr/cantSplit"));

            t.Rows[0].BreakAcrossPages = false;
            Assert.False(t.Rows[0].BreakAcrossPages);
            Assert.Single(t.Rows[0].Xml.RemoveNamespaces().XPathSelectElements("//trPr/cantSplit"));
        }

        [Fact]
        public void MergeCellsAffectGridSpan()
        {
            Table t = new Table(new[,] {{ "1", "2", "3", "4"}});

            // Merge cells 2-4
            t.Rows[0].MergeCells(1,3);

            Assert.Equal(3, t.Rows[0].Cells[1].GridSpan);
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
            Assert.Null(t.PackagePart);

            var doc = Document.Create();
            doc.AddTable(t);
            Assert.NotNull(t.PackagePart);

            var t2 = doc.AddTable(new Table(1, 1));
            Assert.NotNull(t2.PackagePart);
            Assert.Equal(t2.PackagePart, t.PackagePart);
            Assert.Equal(2, doc.Tables.Count());
        }
    }
}
