using System.Drawing;
using DXPlus.Internal;
using Xunit;

namespace DXPlus.Tests
{
    public class ExtensionTests
    {
        [Fact]
        public void ToColorTests()
        {
            Assert.Null(string.Empty.ToColor());
            Assert.Equal(Color.Red.ToArgb(), "ff0000".ToColor()?.ToArgb());
            Assert.Null(((string)null).ToColor());
            Assert.Null("X".ToColor());
            Assert.Equal(Color.Blue.ToArgb(), "ff".ToColor()?.ToArgb());
            Assert.Equal(Color.Red.ToArgb(), "#FF0000".ToColor()?.ToArgb());

            Assert.Equal(Color.Empty, "auto".ToColor());
            Assert.Equal(Color.Empty, "AUTO".ToColor());
            Assert.Equal(Color.Empty, "Auto ".ToColor());
        }

        [Fact]
        public void ToByteTests()
        {
            Assert.Equal((byte)0xd0, "D0".ToByte());
            Assert.NotEqual((byte)0xd0, "12".ToByte());
            Assert.NotEqual((byte)0xff, "FFF".ToByte());
        }

    }
}
