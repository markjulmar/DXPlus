﻿using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DXPlus
{
    /// <summary>
    /// Extension methods for Containers
    /// </summary>
    public static class ContainerHelpers
    {
        /// <summary>
        /// Add an empty paragraph at the end of this container.
        /// </summary>
        /// <returns>Newly added paragraph</returns>
        public static Paragraph AddParagraph(this IContainer container) => container.AddParagraph(string.Empty);

        /// <summary>
        /// Add a new paragraph with the given text at the end of this container.
        /// </summary>
        /// <returns>Newly added paragraph</returns>
        public static Paragraph AddParagraph(this IContainer container, string text) => container.AddParagraph(text, null);

        /// <summary>
        /// Insert a paragraph with the given text at the specified paragraph index.
        /// </summary>
        /// <param name="container">Container to insert into</param>
        /// <param name="index">Index to insert new paragraph at</param>
        /// <param name="text">Text for new paragraph</param>
        /// <returns>Newly added paragraph</returns>
        public static Paragraph InsertParagraph(this IContainer container, int index, string text) => container.InsertParagraph(index, text, null);

        /// <summary>
        /// Add a new equation using the specified text at the end of this container.
        /// </summary>
        /// <param name="container">Container to add equation to</param>
        /// <param name="equation">Equation</param>
        /// <returns>Newly added paragraph</returns>
        public static Paragraph AddEquation(this IContainer container, string equation) => container.AddParagraph().AppendEquation(equation);

        /// <summary>
        /// Add a bookmark to the end of this container.
        /// </summary>
        /// <param name="container">Container</param>
        /// <param name="bookmarkName">Name of the new bookmark</param>
        /// <returns>Newly added paragraph</returns>
        public static Paragraph AddBookmark(this IContainer container, string bookmarkName) => container.AddParagraph().AppendBookmark(bookmarkName);

        /// <summary>
        /// Add a new table to the end of this container
        /// </summary>
        /// <param name="container">Container owner</param>
        /// <param name="rows">Rows to add</param>
        /// <param name="columns">Columns to add</param>
        /// <returns></returns>
        public static Table AddTable(this IContainer container, int rows, int columns) => container.AddTable(new Table(rows, columns));

        /// <summary>
        /// Find all occurrences of a string in the paragraph
        /// </summary>
        /// <param name="container"></param>
        /// <param name="text"></param>
        /// <param name="ignoreCase"></param>
        /// <returns></returns>
        public static IEnumerable<int> FindAll(this IContainer container, string text, bool ignoreCase = false)
        {
            return from p in container.Paragraphs
                from index in p.FindAll(text, ignoreCase)
                select index + p.StartIndex;
        }

        /// <summary>
        /// Find all unique instances of the given Regex Pattern,
        /// returning the list of the unique strings found
        /// </summary>
        /// <param name="container"></param>
        /// <param name="regex">Pattern to search for</param>
        /// <returns>Index and matched strings</returns>
        public static IEnumerable<(int index, string text)> FindPattern(this IContainer container, Regex regex)
        {
            foreach (var p in container.Paragraphs)
            {
                foreach (var (index, text) in p.FindPattern(regex))
                {
                    yield return (index: index + p.StartIndex, text);
                }
            }
        }


    }
}
