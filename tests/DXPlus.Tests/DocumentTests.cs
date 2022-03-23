using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace DXPlus.Tests
{
    public class DocumentTests
    {
        [Fact]
        public void DocPropertyPreservesStartEndSpaces()
        {
            var doc = Document.Create();

            string text = " This is a test ";

            doc.Properties.Title = text;
            Assert.Equal(text, doc.Properties.Title);
        }

        [Fact]
        public void ReplaceDocPropertyChangesValue()
        {
            var doc = Document.Create();

            Assert.NotNull(doc.Properties.Creator);
            Assert.Equal(Environment.UserName, doc.Properties.Creator);
            doc.Properties.Creator = "tom";
            Assert.NotNull(doc.Properties.Creator);
            Assert.Equal("tom", doc.Properties.Creator);
        }

        [Fact]
        public void SetLastModifiedByIsValid()
        {
            var doc = Document.Create();
            Assert.NotNull(doc.Properties.LastSavedBy);
            doc.Properties.LastSavedBy = "tom";
            Assert.Equal("tom", doc.Properties.LastSavedBy);
        }

        [Fact]
        public void InsertedDocPropertyPreservesStartEndSpaces()
        {
            var doc = Document.Create();

            string text = " is a test ";
            doc.Properties.Title = text;

            var p = doc.Add("This")
                .AddDocumentPropertyField(DocumentPropertyName.Title)
                .AddText("of the emergency broadcast system.");
            
            Assert.Equal("This is a test of the emergency broadcast system.", p.Text);
        }
        
        [Fact]
        public void CustomPropertyPreservesStartEndSpaces()
        {
            var doc = Document.Create();

            string text = " This is a test ";
            
            doc.CustomProperties.Add("NewProp", text);

            Assert.True(doc.CustomProperties.TryGetValue("NewProp", out var cp));
            Assert.Equal(text, cp.Value);
        }
        
        [Fact]
        public void InsertedCustomPropertyPreservesStartEndSpaces()
        {
            var doc = Document.Create();
            string text = " is a test ";
            doc.CustomProperties.Add("NewProp", text);

            var p = doc.Add("This")
                .AddCustomPropertyField("NewProp")
                .AddText("of the emergency broadcast system.");
            
            Assert.Equal("This is a test of the emergency broadcast system.", p.Text);
        }

        [Fact]
        public void CannotDeleteRequiredCoreProperties()
        {
            using var doc = Document.Create();

            doc.Properties.Title = null;
            Assert.NotNull(doc.Properties.Title);
            Assert.Empty(doc.Properties.Title);
        }

        [Fact]
        public void AddDocumentPropertyUpdatesCoreXml()
        {
            string text = "Sample Category";
            var filename = Path.Combine(Path.GetTempPath(), Path.GetTempFileName());

            try
            {
                using (var doc = Document.Create(filename))
                {
                    doc.Properties.Category = text;
                    doc.Save();
                }

                using (var doc = Document.Load(filename))
                {
                    Assert.Equal(text, doc.Properties.Category);
                }
            }
            finally
            {
                File.Delete(filename);
            }

        }

        [Fact]
        public void UpdatedDocPropertySetsComplexField()
        {
            var doc = Document.Create();
            string text = "one";
            doc.Properties.Title = text;

            var p = doc.Add("This is number ")
                .AddDocumentPropertyField(DocumentPropertyName.Title)
                .AddText(".");
            
            Assert.Equal("This is number one.", p.Text);

            doc.Properties.Title = "two";
            Assert.Equal("This is number two.", p.Text);
        }
    }
}