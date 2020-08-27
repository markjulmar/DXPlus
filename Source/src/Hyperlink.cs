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
    public class Hyperlink : DocXBase
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
        public void Remove() => Xml.Remove();

        /// <summary>
        /// Change the Text of a Hyperlink.
        /// </summary>
        public string Text
        {
            get => text;

            set
            {
                var rPr = new XElement(Name.RunProperties,
                                new XElement(Namespace.Main + "rStyle",
                                    new XAttribute(Name.MainVal, "Hyperlink")));

                // Format and add the new text.
                var newRuns = HelperFunctions.FormatInput(value, rPr);
                if (type == 0)
                {
                    Xml.LocalNameElements("r").Remove();
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

        /// <summary>
        /// Create a new hyperlink to add to a document
        /// </summary>
        /// <param name="text">Link text</param>
        /// <param name="uri">Link URI</param>
        public Hyperlink(string text, Uri uri)
        {
            Xml = new XElement(Namespace.Main + "hyperlink",
                new XAttribute(Namespace.RelatedDoc + "id", string.Empty),
                new XAttribute(Namespace.Main + "history", "1"),
                new XElement(Name.Run,
                    new XElement(Name.RunProperties,
                        new XElement(Namespace.Main + "rStyle",
                            new XAttribute(Name.MainVal, "Hyperlink"))),
                    new XElement(Name.Text, text))
            );

            type = 0;
            Uri = uri;
            this.text = text;
            Id = string.Empty;
        }

        /// <summary>
        /// Get or create a relationship link to a hyperlink
        /// </summary>
        /// <returns>Relationship id</returns>
        internal string GetOrCreateRelationship()
        {
            string baseUri = Uri.OriginalString;

            // Search for a relationship with a TargetUri that points at this Image.
            string id = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/hyperlink")
                .Where(r => r.TargetUri.OriginalString == baseUri)
                .Select(r => r.Id)
                .SingleOrDefault();

            // No id yet, create one.
            if (id == null)
            {
                // Check to see if a relationship for this Hyperlink exists and create it if not.
                var pr = PackagePart.CreateRelationship(Uri, TargetMode.External,
                    $"{Namespace.RelatedDoc.NamespaceName}/hyperlink");
                id = pr.Id;
            }

            Id = id;
            Xml.SetAttributeValue(Namespace.RelatedDoc + "id", id);

            return id;
        }

        /// <summary>
        /// Internal constructor used when creating hyperlinks out of the document
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML fragment representing hyperlink</param>
        internal Hyperlink(IDocument document, XElement xml) : base(document, xml)
        {
            type = 0;
            Id = xml.AttributeValue(Namespace.RelatedDoc + "id");
            text = HelperFunctions.GetTextRecursive(xml).ToString();
        }

        /// <summary>
        /// Internal constructor used when creating hyperlinks out of the document
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="instrText">Text with field codes</param>
        /// <param name="runs">Text runs making up this hyperlink</param>
        internal Hyperlink(IDocument document, XElement instrText, List<XElement> runs) : base(document, null)
        {
            type = 1;
            this.instrText = instrText;
            this.runs = runs;

            int start = instrText.Value.IndexOf("HYPERLINK \"", StringComparison.Ordinal) + "HYPERLINK \"".Length;
            int end = instrText.Value.IndexOf("\"", start, StringComparison.Ordinal);
            if (start != -1 && end != -1)
            {
                Uri = new Uri(instrText.Value[start..end], UriKind.Absolute);
                text = HelperFunctions.GetTextRecursive(new XElement(Namespace.Main + "temp", runs)).ToString();
            }
        }

        internal static IEnumerable<Hyperlink> Enumerate(DocXBase owner)
        {
            foreach (var he in owner.Xml.Descendants()
                                .Where(h => h.Name.LocalName == "hyperlink" 
                                    || h.Name.LocalName == "instrText").ToList())
            {
                if (he.Name.LocalName == "hyperlink")
                {
                    yield return new Hyperlink(owner.Document, he) {PackagePart = owner.PackagePart};
                }
                else
                {
                    // Find the parent run, no matter how deeply nested we are.
                    var e = he;
                    while (e != null && e.Name.LocalName != "r")
                        e = e.Parent;

                    if (e == null)
                        throw new Exception("Failed to locate the parent in a run.");

                    // Take every element until we reach w:fldCharType="end"
                    var hyperLinkRuns = new List<XElement>();
                    foreach (var run in e.ElementsAfterSelf(Name.Run))
                    {
                        // Add this run to the list.
                        hyperLinkRuns.Add(run);

                        var fldChar = run.Descendants(Namespace.Main + "fldChar").SingleOrDefault();
                        if (fldChar != null)
                        {
                            var fldCharType = fldChar.Attribute(Namespace.Main + "fldCharType");
                            if (fldCharType?.Value.Equals("end", StringComparison.CurrentCultureIgnoreCase) == true)
                            {
                                yield return new Hyperlink(owner.Document, he, hyperLinkRuns) {PackagePart = owner.PackagePart};
                                break;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected override void OnDocumentOwnerChanged(IDocument previousValue, IDocument newValue)
        {
            base.OnDocumentOwnerChanged(previousValue, newValue);
            // Add the hyperlink styles to the document if missing.
            (newValue as DocX)?.AddHyperlinkStyleIfNotPresent();
        }
    }
}