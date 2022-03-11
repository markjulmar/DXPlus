using System;
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
            
            doc.SetPropertyValue(DocumentPropertyName.Title, text);
            Assert.Equal(text, doc.DocumentProperties[DocumentPropertyName.Title]);
        }

        [Fact]
        public void ReplaceDocPropertyChangesValue()
        {
            var doc = Document.Create();

            Assert.True(doc.DocumentProperties.TryGetValue(DocumentPropertyName.Creator, out string lsb));
            Assert.Equal(Environment.UserName, lsb);
            doc.SetPropertyValue(DocumentPropertyName.Creator, "tom");
            Assert.True(doc.DocumentProperties.TryGetValue(DocumentPropertyName.Creator, out string lsb2));
            Assert.Equal("tom", lsb2);
        }

        [Fact]
        public void SetLastModifiedByIsValid()
        {
            var doc = Document.Create();

            Assert.True(doc.DocumentProperties.TryGetValue(DocumentPropertyName.LastSavedBy, out string lsb));

            doc.SetPropertyValue(DocumentPropertyName.LastSavedBy, "tom");
            Assert.True(doc.DocumentProperties.TryGetValue(DocumentPropertyName.LastSavedBy, out string lsb2));
            Assert.Equal("tom", lsb2);
        }

        [Fact]
        public void InsertedDocPropertyPreservesStartEndSpaces()
        {
            var doc = Document.Create();
            string text = " is a test ";
            doc.SetPropertyValue(DocumentPropertyName.Title, text);

            var p = doc.Add("This")
                .AddDocumentPropertyField(DocumentPropertyName.Title)
                .Append("of the emergency broadcast system.");
            
            Assert.Equal("This is a test of the emergency broadcast system.", p.Text);
        }
        
        [Fact]
        public void CustomPropertyPreservesStartEndSpaces()
        {
            var doc = Document.Create();

            string text = " This is a test ";
            
            doc.AddCustomProperty("NewProp", text);
            Assert.Equal(text, doc.CustomProperties["NewProp"]);
        }
        
        [Fact]
        public void InsertedCustomPropertyPreservesStartEndSpaces()
        {
            var doc = Document.Create();
            string text = " is a test ";
            doc.AddCustomProperty("NewProp", text);

            var p = doc.Add("This")
                .AddCustomPropertyField("NewProp")
                .Append("of the emergency broadcast system.");
            
            Assert.Equal("This is a test of the emergency broadcast system.", p.Text);
        }

        [Fact]
        public void UpdatedDocPropertySetsComplexField()
        {
            var doc = Document.Create();
            string text = "one";
            doc.SetPropertyValue(DocumentPropertyName.Title, text);

            var p = doc.Add("This is number ")
                .AddDocumentPropertyField(DocumentPropertyName.Title)
                .Append(".");
            
            Assert.Equal("This is number one.", p.Text);
            
            doc.SetPropertyValue(DocumentPropertyName.Title, "two");
            Assert.Equal("This is number two.", p.Text);
        }
    }
}