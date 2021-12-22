using System;
using System.IO;
using System.Linq;
using DXPlus;

namespace Tester
{
    class Program
    {
        public static void Main(string[] args)
        {
            CreateTableWithList();
        }

        static void CreateTableWithList()
        {
            string fn = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.docx");
            using var doc = Document.Create(fn);

            int rows = 2;
            int columns = 2;

            var documentTable = doc.AddTable(rows: rows, columns: columns);
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            documentTable.Design = TableDesign.TableNormal;

            AddListToCell(
                documentTable.Rows.ElementAt(1).Cells[0], nd,
                new[] { "Item 1", "Item 2", "Item 3", "Item 4" });

            AddListToCell(
                documentTable.Rows.ElementAt(1).Cells[1], nd,
                new[] { "Item 5", "Item 6", "Item 7", "Item 8" });

            documentTable.AutoFit = true;

            var t = doc.AddTable(3, 3);

            doc.Save();

            Console.WriteLine("Wrote document");
        }

        private static void AddListToCell(TableCell cell, NumberingDefinition nd, string[] terms)
        {
            var cellParagraph = cell.Paragraphs.First();
            for (int i = 0; i < terms.Length; i++)
            {
                string text = terms[i];
                if (i == 0)
                    cellParagraph.ListStyle(nd, 0);
                else
                    cellParagraph = cellParagraph.AddParagraph().ListStyle(nd,0);
                cellParagraph.SetText(text);
            }
        }
    }
}
