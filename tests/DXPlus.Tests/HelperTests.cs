using DXPlus.Helpers;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class HelperTests
    {
        [Fact]
        public void BasicHexNumberTestsReturnTrueFalse()
        {
            string value = 512.ToString("X2");
            Assert.True(HelperFunctions.IsValidHexNumber(value));
            Assert.True(HelperFunctions.IsValidHexNumber("FFFFFFFF"));
            Assert.True(HelperFunctions.IsValidHexNumber("0"));
            Assert.False(HelperFunctions.IsValidHexNumber(null));
            Assert.False(HelperFunctions.IsValidHexNumber(""));
            Assert.False(HelperFunctions.IsValidHexNumber(" "));
            Assert.False(HelperFunctions.IsValidHexNumber("test"));
        }

        [Fact]
        public void EnumToHexReturnsZeroFilledValue()
        {
            var val = TableConditionalFormatting.FirstColumn
                      | TableConditionalFormatting.FirstRow;

            string s = val.ToHex(4);
            Assert.Equal("00A0", s);
        }

        [Fact]
        public void EnumToHexWithNoLengthReturnsUnfilledValue()
        {
            var val = TableConditionalFormatting.FirstColumn
                      | TableConditionalFormatting.FirstRow;

            string s = val.ToHex();
            Assert.Equal("A0", s);
        }

        [Fact]
        public void GetAttributeValueReturnValueOrNull()
        {
            var xml = XElement.Parse(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <tbl>
                       <tcPr>
                          <tcMar test='1'>
                            <bottom w=""10"" test='2' />
                        </tcMar>
                        </tcPr>
                    </tbl>");

            string val = xml.AttributeValue("tcPr", "tcMar", "bottom", "w");
            Assert.Equal("10", val);

            val = xml.AttributeValue("tcPr", "tcMar", "top", "w");
            Assert.Null(val);
        }

        [Fact]
        public void GetAttributeValueWorksWithNamespaces()
        {
            XNamespace Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var xml = XDocument.Parse(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                        <w:document xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
                                    xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                          <w:body>
                            <w:sectPr>
                              <w:pgSz w:h=""15840"" w:orient=""portrait"" w:w=""12240"" />
                              <w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""720"" w:footer=""720"" w:gutter=""0"" />
                              <w:cols w:space=""720"" />
                              <w:docGrid w:linePitch=""360"" />
                              <w:headerReference w:type=""first"" r:id=""R89b5d4ad56c64367"" />
                              <w:titlePg></w:titlePg>
                            </w:sectPr>
                          </w:body>
                        </w:document>");

            string val = xml.Root.AttributeValue(Main + "body", Main + "sectPr", Main + "cols", Main + "space");
            Assert.Equal("720", val);
            val = xml.Root.AttributeValue("body", "sectPr", Main + "cols", Main + "space");
            Assert.Null(val);
        }

        [Fact]
        public void GetElementReturnValueOrNull()
        {
            var xml = XElement.Parse(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <tbl>
                       <tcPr>
                          <tcMar test='1'>
                            <bottom w=""10"" test='2' />
                        </tcMar>
                        </tcPr>
                    </tbl>");

            var val = xml.Element("tcPr", "tcMar", "bottom");
            Assert.Equal("bottom", val.Name.LocalName);

            val = xml.Element("tcPr", "tcMgn", "bottom");
            Assert.Null(val);
        }

        [Fact]
        public void GetElementWorksWithNamespaces()
        {
            XNamespace Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var xml = XDocument.Parse(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                        <w:document xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
                                    xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                          <w:body>
                            <w:sectPr>
                              <w:pgSz w:h=""15840"" w:orient=""portrait"" w:w=""12240"" />
                              <w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""720"" w:footer=""720"" w:gutter=""0"" />
                              <w:cols w:space=""720"" />
                              <w:docGrid w:linePitch=""360"" />
                              <w:headerReference w:type=""first"" r:id=""R89b5d4ad56c64367"" />
                              <w:titlePg></w:titlePg>
                            </w:sectPr>
                          </w:body>
                        </w:document>");

            var val = xml.Root.Element(Main + "body", Main + "sectPr", Main + "cols");
            Assert.Equal("cols", val?.Name.LocalName);
            val = xml.Element("body", Main + "sectPr", Main + "cols");
            Assert.Null(val);
        }

        [Fact]
        public void GetElementsReturnValueOrEmpty()
        {
            var xml = XElement.Parse(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <tbl>
                        <row>
                           <line>
                              <entry />
                              <entry />
                              <entry />
                              <entry />
                              <entry />
                            </line>
                        </row>
                    </tbl>");

            var vals = xml.Elements("row", "line");
            Assert.Equal(5, vals.Count());
            vals = xml.Elements("tcPr", "tcMgn", "bottom");
            Assert.Empty(vals);
        }

        [Fact]
        public void SetAttributeValueCreatesPath()
        {
            var xml = XElement.Parse(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                    <tbl>
                    </tbl>");

            var attr = xml.SetAttributeValue("tbl", "line", "entry", "value", 10);
            Assert.NotNull(attr);
            Assert.Equal("value", attr.Name.LocalName);
            Assert.Equal("10", attr.Value);
        }

        [Fact]
        public void NormalizeOrdersElements()
        {
            XElement e1 = XElement.Parse(
                @"<root>
                  <a>
                    <margin val='2'></margin>
                    <value val='1' />
                  </a>
                  <a>
                    <value val='2' />
                    <margin val='3'></margin>
                  </a> 
                  <a>
                    <value val='3' />
                    <margin val='4'></margin>
                  </a>
            </root>");

            XElement e2 = XElement.Parse(
                @"<root>
                  <a>
                    <value val='1' />
                    <margin val='2'></margin>
                  </a>
                  <a>
                    <margin val='3'></margin>
                    <value val='2' />
                  </a>
                  <a>
                    <value val='3' />
                    <margin val='4'></margin>
                  </a> 
            </root>");

            var ne = e1.Normalize();
            var ne2 = e2.Normalize();
            Assert.True(XNode.DeepEquals(ne, ne2));
        }

        [Fact]
        public void FindLastUsedIdReturnsZeroWhenNone()
        {
            var doc = XDocument.Parse(
                @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                <root xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                </root>");

            long id = HelperFunctions.FindLastUsedDocId(doc);
            Assert.Equal(0, id);
        }

        [Fact]
        public void FindLastUsedIdReturnsHighestValue()
        {
            var doc = XDocument.Parse(
                @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                <w:root xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                    <w:item>
                        <w:bookmarkStart w:id=""20"" />
                    </w:item>
                    <w:item>
                        <w:bookmarkStart w:id=""2"" />
                    </w:item>
                    <w:item>
                        <w:bookmarkStart w:id=""15"">
                            <wp:docPr id=""16"" />
                        </w:bookmarkStart>
                    </w:item>
                </w:root>");

            Assert.Equal(20, HelperFunctions.FindLastUsedDocId(doc));
            doc.XPathSelectElement("//w:bookmarkStart[@w:id='20']", Namespace.NamespaceManager()).Remove();
            Assert.Equal(16, HelperFunctions.FindLastUsedDocId(doc));
        }
    }
}
