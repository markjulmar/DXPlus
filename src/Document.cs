using System.IO;

namespace DXPlus
{
    /// <summary>
    /// Public facing class to create or load a document.
    /// </summary>
    public static class Document
    {
        /// <summary>
        /// Load a Word document from a file.
        /// </summary>
        /// <param name="filename">Filename</param>
        /// <returns>Loaded document</returns>
        public static IDocument Load(string filename)
        {
            return DocX.Load(filename);
        }

        /// <summary>
        /// Load a Word document from a stream
        /// </summary>
        /// <param name="stream">Stream</param>
        /// <returns>Loaded document</returns>
        public static IDocument Load(Stream stream)
        {
            return DocX.Load(stream);
        }

        /// <summary>
        /// Create a new Word document
        /// </summary>
        /// <param name="filename">Optional filename</param>
        /// <returns>New document</returns>
        public static IDocument Create(string filename = null)
        {
            return DocX.Create(filename, DocumentTypes.Document);
        }

        /// <summary>
        /// Create a new Word document template
        /// </summary>
        /// <param name="filename">Optional filename</param>
        /// <returns>New template</returns>
        public static IDocument CreateTemplate(string filename = null)
        {
            return DocX.Create(filename, DocumentTypes.Template);
        }
    }
}