﻿using DXPlus.Helpers;
using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a block of content in the document. This can be a Paragraph, Numbered list, or Table.
    /// </summary>
    public abstract class Block : DocXElement
    {
        private BlockContainer owner;

        /// <summary>
        /// Default constructor used with Lists/Tables
        /// </summary>
        internal Block()
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document"></param>
        /// <param name="xml"></param>
        protected Block(IDocument document, XElement xml) : base(document, xml)
        {
        }

        /// <summary>
        /// Add a page break after the current element.
        /// </summary>
        public void AddPageBreak() => Xml.AddAfterSelf(HelperFunctions.PageBreak());

        /// <summary>
        /// Insert a page break before the current element.
        /// </summary>
        public void InsertPageBreakBefore() => Xml.AddBeforeSelf(HelperFunctions.PageBreak());

        /// <summary>
        /// Add a new paragraph after the current element.
        /// </summary>
        /// <param name="paragraph">Paragraph to insert</param>
        public void AddParagraph(Paragraph paragraph)
        {
            if (paragraph.BlockContainer != null)
                throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

            Xml.AddAfterSelf(paragraph.Xml);

            if (owner != null)
            {
                paragraph.BlockContainer = owner;
                paragraph.SetStartIndex(owner.Paragraphs.Single(p => p.Id == paragraph.Id).StartIndex);
            }
        }

        /// <summary>
        /// Add a paragraph after the current element using the passed text
        /// </summary>
        /// <param name="text">Text for new paragraph</param>
        /// <param name="formatting">Formatting for the paragraph</param>
        /// <returns>Newly created paragraph</returns>
        public Paragraph AddParagraph(string text, Formatting formatting)
        {
            var paragraph = new Paragraph(Document, Paragraph.Create(text, formatting), -1);
            AddParagraph(paragraph);
            return paragraph;
        }

        /// <summary>
        /// Insert a paragraph before the current element
        /// </summary>
        /// <param name="paragraph"></param>
        public void InsertParagraphBefore(Paragraph paragraph)
        {
            if (paragraph.BlockContainer != null)
                throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

            Xml.AddBeforeSelf(paragraph.Xml);
            
            if (owner != null)
            {
                paragraph.BlockContainer = owner;
                paragraph.SetStartIndex(owner.Paragraphs.Single(p => p.Id == paragraph.Id).StartIndex);
            }
        }

        /// <summary>
        /// Insert a paragraph before the current element
        /// </summary>
        /// <param name="text">Text to use for new paragraph</param>
        /// <param name="formatting">Formatting to use</param>
        /// <returns></returns>
        public Paragraph InsertParagraphBefore(string text, Formatting formatting)
        {
            var paragraph = new Paragraph(Document, Paragraph.Create(text, formatting), -1);
            InsertParagraphBefore(paragraph);
            return paragraph;
        }

        /// <summary>
        /// Add a new table after this container
        /// </summary>
        /// <param name="table">Table to add</param>
        public void AddTable(Table table)
        {
            if (table.BlockContainer != null)
                throw new ArgumentException("Cannot add table multiple times.", nameof(table));

            table.BlockContainer = this.BlockContainer;
            Xml.AddAfterSelf(table.Xml);
        }

        /// <summary>
        /// Insert a table before this element
        /// </summary>
        /// <param name="table"></param>
        public void InsertTableBefore(Table table)
        {
            if (table.BlockContainer != null)
                throw new ArgumentException("Cannot add table multiple times.", nameof(table));

            table.BlockContainer = BlockContainer;
            Xml.AddBeforeSelf(table.Xml);
        }

        /// <summary>
        /// Set the container owner for this element.
        /// </summary>
        internal BlockContainer BlockContainer
        {
            get => owner;
            set
            {
                if (owner == value)
                    return;

                if (owner != null)
                    OnRemovedFromContainer(owner);

                owner = value;
                PackagePart = owner?.PackagePart;
                Document = owner?.Document;

                if (owner != null)
                    OnAddedToContainer(owner);
            }
        }

        /// <summary>
        /// Invoked when this element is added to a container.
        /// </summary>
        /// <param name="blockContainer">Current Container owner</param>
        protected virtual void OnAddedToContainer(BlockContainer blockContainer)
        {
        }

        /// <summary>
        /// Invoked when this element is removed from a container.
        /// </summary>
        /// <param name="blockContainer">Current Container owner</param>
        protected virtual void OnRemovedFromContainer(BlockContainer blockContainer)
        {
        }
    }
}