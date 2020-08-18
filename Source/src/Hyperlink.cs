using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

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
                var rPr = new XElement(DocxNamespace.Main + "rPr",
                                new XElement(DocxNamespace.Main + "rStyle",
                                    new XAttribute(DocxNamespace.Main + "val", "Hyperlink")));

                // Format and add the new text.
                var newRuns = HelperFunctions.FormatInput(value, rPr);
                if (type == 0)
                {
                    // Get all the runs in this Text.
                    var runs = Xml.LocalNameElements("r").ToList();
                    for (int i = 0; i < runs.Count; i++)
                    {
                        runs.Remove();
                    }

                    Xml.Add(newRuns);
                }
                else
                {
                    var separate = XElement.Parse(@"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='separate'/>
                    </w:r>");

                    var end = XElement.Parse(@"
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
            get => type == 0 && !string.IsNullOrEmpty(Id) 
                ? PackagePart.GetRelationship(Id).TargetUri 
                : uri;

            set
            {
                uri = value;

                if (type == 0)
                {
                    if (!string.IsNullOrEmpty(Id))
                    {
                        var r = PackagePart.GetRelationship(Id);

                        // Get all of the information about this relationship.
                        var targetMode = r.TargetMode;
                        string relationshipType = r.RelationshipType;
                        string id = r.Id;

                        PackagePart.DeleteRelationship(id);
                        PackagePart.CreateRelationship(value, targetMode, relationshipType, id);
                    }
                }
                else
                {
                    instrText.Value = $"HYPERLINK \"{value}\"";
                }
            }
        }

        internal Hyperlink(DocX document, XElement data, PackagePart packagePart) : base(document, data)
        {
            type = 0;
            Id = data.AttributeValue(DocxNamespace.RelatedDoc + "id");
            text = HelperFunctions.GetTextRecursive(data).ToString();
            PackagePart = packagePart;
        }

        internal Hyperlink(DocX document, XElement instrText, List<XElement> runs) : base(document, null)
        {
            type = 1;
            this.instrText = instrText;
            this.runs = runs;

            int start = instrText.Value.IndexOf("HYPERLINK \"", StringComparison.Ordinal) + "HYPERLINK \"".Length;
            int end = instrText.Value.IndexOf("\"", start, StringComparison.Ordinal);
            if (start != -1 && end != -1)
            {
                Uri = new Uri(instrText.Value[start..end], UriKind.Absolute);
                text = HelperFunctions.GetTextRecursive(new XElement(DocxNamespace.Main + "temp", runs)).ToString();
            }
        }
    }
}