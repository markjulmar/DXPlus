using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    public class Header : Container
    {
        public bool PageNumbers
        {
            get => false;

            set
            {
                var rsid = Document.RevisionId;

                XElement e = XElement.Parse
                ($@"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:sdtPr>
                      <w:id w:val='157571950' />
                      <w:docPartObj>
                        <w:docPartGallery w:val='Page Numbers (Top of Page)' />
                        <w:docPartUnique />
                      </w:docPartObj>
                    </w:sdtPr>
                    <w:sdtContent>
                      <w:p w:rsidR='{rsid}' w:rsidRDefault='{rsid}'>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                          <w:jc w:val='center' />
                        </w:pPr>
                        <w:fldSimple w:instr=' PAGE \* MERGEFORMAT'>
                          <w:r>
                            <w:rPr>
                              <w:noProof />
                            </w:rPr>
                            <w:t>1</w:t>
                          </w:r>
                        </w:fldSimple>
                      </w:p>
                    </w:sdtContent>
                  </w:sdt>"
               );

                Xml.AddFirst(e);
            }
        }

        public string Id { get; set; }

        internal Header(DocX document, XElement xml, PackagePart mainPart) : base(document, xml)
        {
            PackagePart = mainPart;
        }

        public IEnumerable<Image> Images => PackagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                                  .Select(i => new Image(Document, i));
    }
}