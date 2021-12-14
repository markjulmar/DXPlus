using System.Collections.Generic;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.TripleColonExtensions
{
    public sealed class TripleColonElement
    {
        public ITripleColonExtensionInfo Extension { get; set; }
        public IDictionary<string, string> RenderProperties { get; set; }
        public bool Closed { get; set; }
        public bool EndingTripleColons { get; set; }
        public IDictionary<string, string> Attributes { get; set; }
        public int Count { get; }

        public TripleColonElement(TripleColonBlock block)
        {
            Extension = block.Extension;
            RenderProperties = block.RenderProperties;
            Closed = block.Closed;
            EndingTripleColons = block.EndingTripleColons;
            Attributes = block.Attributes;
            Count = block.Count;
        }

        public TripleColonElement(TripleColonInline inline)
        {
            Extension = inline.Extension;
            RenderProperties = inline.RenderProperties;
            Closed = inline.Closed;
            EndingTripleColons = inline.EndingTripleColons;
            Attributes = inline.Attributes;
            Count = inline.Count;
        }
    }
}