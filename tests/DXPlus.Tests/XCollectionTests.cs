using System;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Internal;
using Xunit;
using static System.Int32;

namespace DXPlus.Tests
{
    public class TestElement : XElementWrapper, IEquatable<TestElement>
    {
        public int Index => Parse(Xml.AttributeValue("num"));
        internal TestElement(XElement xe) { Xml = xe; }
        public TestElement(int index) { Xml = new XElement("value", new XAttribute("num", index)); }
        public bool Equals(TestElement other) => other is not null && (ReferenceEquals(this, other) || Index==other.Index);
    }

    public class TestEC : XElementCollection<TestElement>
    {
        public XElement root;

        public TestEC(XElement parent) : base(parent, "values", "value", xe => new TestElement(xe))
        {
            root = parent;
        }
    }

    public class XCollectionTests
    {
        private TestEC Create()
        {
            var parent = new XElement("test", new XElement("values",
                new XElement("value", new XAttribute("num", 1)),
                new XElement("value", new XAttribute("num", 3)),
                new XElement("value", new XAttribute("num", 5)),
                new XElement("value", new XAttribute("num", 6)),
                new XElement("value", new XAttribute("num", 7))));

            return new TestEC(parent);
        }

        [Fact]
        public void ParsesXmlIntoChildren()
        {
            var t = Create();

            Assert.NotEmpty(t);
            Assert.Equal(5, t.Count);
            
            Assert.Equal(1, t[0].Index);
            Assert.Equal(3, t[1].Index);
            Assert.Equal(5, t[2].Index);
            Assert.Equal(6, t[3].Index);
            Assert.Equal(7, t[4].Index);

            Assert.Throws<ArgumentOutOfRangeException>(() => t[10]);
            Assert.Throws<ArgumentOutOfRangeException>(() => t[-1]);
        }

        [Fact]
        public void CanCopyIntoArray()
        {
            var t = Create();

            var array = new TestElement[8];
            t.CopyTo(array,2);
            Assert.Null(array[0]);
            Assert.Null(array[1]);
            Assert.NotNull(array[2]);
            Assert.NotNull(array[6]);
            Assert.Null(array[7]);

            var item = array[2];
            Assert.Equal(t[0], item);
        }

        [Fact]
        public void CanCreateNewItem()
        {
            var t = Create();

            Assert.Equal(5, t.Count);
            t.Add(new TestElement(10));
            Assert.Equal(10, t.Last().Index);
            Assert.Equal(6, t.Count);

            Assert.Equal(1, t[0].Index);
            Assert.Equal(3, t[1].Index);
            Assert.Equal(5, t[2].Index);
            Assert.Equal(6, t[3].Index);
            Assert.Equal(7, t[4].Index);
        }

        [Fact]
        public void CanReplaceItems()
        {
            var t = Create();

            t[1] = new TestElement(10);
            Assert.Equal(10, t[1].Index);
            Assert.Equal(5, t.Count);

            Assert.Equal(1, t[0].Index);
            Assert.Equal(10, t[1].Index);
            Assert.Equal(5, t[2].Index);
            Assert.Equal(6, t[3].Index);
            Assert.Equal(7, t[4].Index);
        }

        [Fact]
        public void CanRemoveItems()
        {
            var t = Create();

            t.RemoveAt(4);
            Assert.Equal(4, t.Count);

            Assert.Equal(1, t[0].Index);
            Assert.Equal(3, t[1].Index);
            Assert.Equal(5, t[2].Index);
            Assert.Equal(6, t[3].Index);
        }

        [Fact]
        public void CanEnumerateItems()
        {
            var t = Create();

            var array = t.ToArray();
            Assert.Equal(5, array.Length);
            for (int i = 0; i < array.Length; i++)
            {
                Assert.Equal(array[i], t[i]);
                Assert.False(ReferenceEquals(array[i], t[i]));
            }
        }

        [Fact]
        public void ClearRemovesAll()
        {
            var t = Create();

            t.Clear();
            Assert.Empty(t);

            Assert.Empty(t.root.Elements());
        }

        [Fact]
        public void RemoveLastItemRemovesTag()
        {
            var t = Create();

            int count = t.Count;
            for (int i = 0; i < count; i++)
            {
                t.RemoveAt(0);
            }

            Assert.Empty(t);
            Assert.Empty(t.root.Elements());

            t.Add(new TestElement(1));
            Assert.Single(t);
            Assert.Equal(1, t[0].Index);
        }

        [Fact]
        public void RemoveAtRemovesXml()
        {
            var t = Create();

            Assert.Equal(5, t[2].Index);
            t.RemoveAt(2);
            Assert.Equal(6, t[2].Index);

            Assert.Throws<ArgumentOutOfRangeException>(() => t.RemoveAt(10));
            Assert.Throws<ArgumentOutOfRangeException>(() => t.RemoveAt(-1));
        }

        [Fact]
        public void IndexOfCanFindItem()
        {
            var t = Create();

            var item = new TestElement(5);
            Assert.Equal(2, t.IndexOf(item));
        }

        [Fact]
        public void CanRemoveItem()
        {
            var t = Create();

            var item = t[2];

            Assert.True(t.Remove(item));
            Assert.False(t.Remove(new TestElement(5)));
        }

        [Fact]
        public void CanInsertAtIndex()
        {
            var t = Create();

            t.Insert(0, new TestElement(0));
            t.Insert(2, new TestElement(2));
            t.Insert(4, new TestElement(4));

            Assert.Equal(8, t.Count);
            Assert.Throws<ArgumentOutOfRangeException>(() => t.Insert(t.Count, new TestElement(10)));

            Assert.Equal(0, t[0].Index);
            Assert.Equal(1, t[1].Index);
            Assert.Equal(2, t[2].Index);
            Assert.Equal(3, t[3].Index);
            Assert.Equal(4, t[4].Index);
            Assert.Equal(5, t[5].Index);
            Assert.Equal(6, t[6].Index);
            Assert.Equal(7, t[7].Index);
        }
    }
}
