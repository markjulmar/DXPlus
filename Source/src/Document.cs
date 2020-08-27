using System.IO;

namespace DXPlus
{
    /// <summary>
    /// Public facing class to create or load a document.
    /// </summary>
    public static class Document
    {
        public static IDocument Load(string filename) => DocX.Load(filename);
        public static IDocument Load(Stream stream) => DocX.Load(stream);
        public static IDocument Create(string filename) => DocX.Create(filename);
        public static IDocument CreateTemplate(string filename) => DocX.Create(filename, DocumentTypes.Template);
    }
}