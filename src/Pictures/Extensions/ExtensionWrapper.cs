using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This wraps an extension list in a drawing/image/picture
    /// </summary>
    public sealed class ExtensionWrapper : ICollection<DrawingExtension>
    {
        private readonly Document document;
        private readonly XElement xmlOwner;
        
        /// <summary>
        /// Wrapper around an extension list (extLst)
        /// </summary>
        /// <param name="document"></param>
        /// <param name="xmlOwner"></param>
        public ExtensionWrapper(Document document, XElement xmlOwner)
        {
            this.xmlOwner = xmlOwner;
            this.document = document;
        }

        /// <summary>
        /// Retrieves the extension list, creates if necessary
        /// </summary>
        /// <param name="create">True to create</param>
        /// <returns>List, or null</returns>
        private XElement GetExtLst(bool create)
        {
            var xml = xmlOwner.Element(Namespace.DrawingMain + "extLst");
            if (xml == null && create)
            {
                xml = new XElement(Namespace.DrawingMain + "extLst");
                xmlOwner.Add(xml);
            }

            return xml;
        }
        
        /// <summary>
        /// Wrapper for known drawing extensions
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        private DrawingExtension WrapExtension(XElement xml)
        {
            if (xml == null)
                return null;
            
            string id = xml.AttributeValue("uri");
            
            if (id == SvgExtension.ExtensionId)
                return new SvgExtension(document, xml);
            if (id == VideoExtension.ExtensionId)
                return new VideoExtension(xml);
            if (id == DecorativeImageExtension.ExtensionId)
                return new DecorativeImageExtension(xml);
            if (id == LocalDpiExtension.ExtensionId)
                return new LocalDpiExtension(xml);
            
            return new DrawingExtension(xml);
        }

        /// <summary>
        /// Method to scan an extension list for a specific extension id
        /// </summary>
        /// <param name="id">ID to look for</param>
        /// <param name="create">True to create it</param>
        /// <returns>XElement of [a:ext]</returns>
        private XElement GetExtension(string id, bool create)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("Value cannot be null or empty.", nameof(id));

            var extension = GetExtLst(false)?.Elements(Namespace.DrawingMain + "ext")
                .SingleOrDefault(e => e.AttributeValue("uri") == id);
            if (extension == null && create)
            {
                extension = new XElement(Namespace.DrawingMain + "ext",
                    new XAttribute("uri", id));
                GetExtLst(true).Add(extension);
            }

            return extension;
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator<DrawingExtension> GetEnumerator()
        {
            var result = GetExtLst(false)?
                .Elements(Namespace.DrawingMain + "ext")
                .Select(WrapExtension) ?? Enumerable.Empty<DrawingExtension>();
            return result.GetEnumerator();

        }

        /// <summary>
        /// Base enumerator
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// Adds an item to the extension collection
        /// </summary>
        public void Add(DrawingExtension extension)
        {
            if (extension == null) 
                throw new ArgumentNullException(nameof(extension));
            
            string id = extension.UriId;
            var ext = GetExtension(id, true);
            ext.ReplaceWith(extension.Xml);
        }

        /// <summary>
        /// Returns whether the given extension ID is in the collection.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public bool Contains(string id) => GetExtension(id, false) != null;

        /// <summary>
        /// Returns the given extension by id.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public DrawingExtension Get(string id) => WrapExtension(GetExtension(id, false));

        /// <summary>
        /// Returns the given extension by id.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        internal T Get<T>() where T : DrawingExtension
        {
            string id = typeof(T).GetField(nameof(SvgExtension.ExtensionId), BindingFlags.Public | BindingFlags.Static)?
                .GetValue(null)?.ToString();
            return !string.IsNullOrEmpty(id) 
                ? (T)WrapExtension(GetExtension(id, false)) : null;
        }
        
        /// <summary>
        /// Removes the given extension by id.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public bool Remove(string id)
        {
            var ext = GetExtension(id, false);
            if (ext != null)
            {
                ext.Remove();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Clear the collection of extensions.
        /// </summary>
        public void Clear()
        {
            GetExtLst(false)?.Remove();
        }

        /// <summary>
        /// Returns whether an extension of the given type is in the collection.
        /// Note this is not an EXACT compare - only the ID is checked.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        bool ICollection<DrawingExtension>.Contains(DrawingExtension item) => 
            item != null && Contains(item.UriId);

        /// <summary>
        /// Used to copy the extensions over. This is not supported.
        /// </summary>
        /// <param name="array"></param>
        /// <param name="arrayIndex"></param>
        /// <exception cref="NotImplementedException"></exception>
        void ICollection<DrawingExtension>.CopyTo(DrawingExtension[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Removes the given extension from the collection.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        bool ICollection<DrawingExtension>.Remove(DrawingExtension item) => 
            item != null && Remove(item.UriId);

        /// <summary>Gets the number of elements contained in the <see cref="T:System.Collections.Generic.ICollection`1" />.</summary>
        /// <returns>The number of elements contained in the <see cref="T:System.Collections.Generic.ICollection`1" />.</returns>
        /// <footer><a href="https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.ICollection-1.Count?view=netcore-5.0">`ICollection.Count` on docs.microsoft.com</a></footer>
        public int Count => GetExtLst(false)?.Elements(Namespace.DrawingMain + "ext").Count()??0;

        /// <summary>Gets a value indicating whether the <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only.</summary>
        /// <returns>
        /// <see langword="true" /> if the <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only; otherwise, <see langword="false" />.</returns>
        /// <footer><a href="https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.ICollection-1.IsReadOnly?view=netcore-5.0">`ICollection.IsReadOnly` on docs.microsoft.com</a></footer>
        bool ICollection<DrawingExtension>.IsReadOnly => false;
    }
}