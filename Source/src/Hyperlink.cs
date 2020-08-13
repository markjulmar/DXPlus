using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a Hyperlink in a document.
    /// </summary>
    public class Hyperlink : DocXElement
    {
        private Uri uri;
        private string text;
        private readonly int type;
        private readonly XElement instrText;
        private readonly List<XElement> runs;

        public string Id { get; set; }

        /// <summary>
        /// Remove a Hyperlink from this Paragraph only.
        /// </summary>
        public void Remove()
        {
            Xml.Remove();
        }

        /// <summary>
        /// Change the Text of a Hyperlink.
        /// </summary>
        public string Text
        {
            get => text;

            set
            {
                XElement rPr = new XElement(DocxNamespace.Main + "rPr",
                        new XElement(DocxNamespace.Main + "rStyle",
                            new XAttribute(DocxNamespace.Main + "val", "Hyperlink")));

                // Format and add the new text.
                List<XElement> newRuns = HelperFunctions.FormatInput(value, rPr);
                if (type == 0)
                {
                    // Get all the runs in this Text.
                    var runs = Xml.LocalNameElements("r").ToList();
                    for (int i = 0; i < runs.Count; i++)
                        runs.Remove();

                    Xml.Add(newRuns);
                }
                else
                {
                    XElement separate = XElement.Parse(@"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='separate'/>
                    </w:r>");

                    XElement end = XElement.Parse(@"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='end' />
                    </w:r>");

                    runs.Last().AddAfterSelf(separate, newRuns, end);
                    runs.ForEach(r => r.Remove());
                }

                text = value;
            }
        }

        /// <summary>
        /// Change the Uri of a Hyperlink.
        /// </summary>
        public Uri Uri
        {
            get
            {
                if (type == 0 && !string.IsNullOrEmpty(Id))
                {
                    PackageRelationship r = packagePart.GetRelationship(Id);
                    return r.TargetUri;
                }

                return uri;
            }

            set
            {
                uri = value;

                if (type == 0)
                {
                    if (!string.IsNullOrEmpty(Id))
                    {
                        PackageRelationship r = packagePart.GetRelationship(Id);

                        // Get all of the information about this relationship.
                        TargetMode r_tm = r.TargetMode;
                        string r_rt = r.RelationshipType;
                        string r_id = r.Id;

                        // Delete the relationship
                        packagePart.DeleteRelationship(r_id);
                        packagePart.CreateRelationship(value, r_tm, r_rt, r_id);
                    }
                }
                else
                {
                    instrText.Value = $"HYPERLINK \"{value}\"";
                }
            }
        }

        internal Hyperlink(DocX document, XElement data, Uri uri) : base(document, data)
        {
            type = 0;
            Id = data.AttributeValue(DocxNamespace.RelatedDoc + "id");
            text = HelperFunctions.GetTextRecursive(data).ToString();
            Uri = uri;
        }

        internal Hyperlink(DocX document, XElement instrText, List<XElement> runs) : base(document, null)
        {
            type = 1;
            this.instrText = instrText;
            this.runs = runs;

            int start = instrText.Value.IndexOf("HYPERLINK \"") + "HYPERLINK \"".Length;
            int end = instrText.Value.IndexOf("\"", start);
            if (start != -1 && end != -1)
            {
                Uri = new Uri(instrText.Value[start..end], UriKind.Absolute);
                text = HelperFunctions.GetTextRecursive(new XElement(DocxNamespace.Main + "temp", runs)).ToString();
            }
        }
    }
}