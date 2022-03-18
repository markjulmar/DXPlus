using System;
using System.Xml.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class CustomPropertyTests
    {
        [Fact]
        public void CanCreateBasicTextCustomProperty()
        {
            var prop = new CustomProperty("test", "text");
            Assert.Equal(CustomPropertyType.Text, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal("text", prop.Value);
        }

        [Fact]
        public void CanCreateBasicDateTimeCustomProperty()
        {
            DateTime expected = new(2021, 1, 2, 3, 4, 5);
            var prop = new CustomProperty("test", expected);
            Assert.Equal(CustomPropertyType.DateTime, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal(expected, prop.As<DateTime>());
        }

        [Fact]
        public void CanCreateBasicIntegerCustomProperty()
        {
            var expected = 2021;
            var prop = new CustomProperty("test", expected);
            Assert.Equal(CustomPropertyType.Integer, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal(expected, prop.As<int>());
            Assert.Null(prop.As<DateTime>());
        }

        [Fact]
        public void CanCreateBasicGuidCustomProperty()
        {
            var expected = Guid.NewGuid();
            var prop = new CustomProperty("test", expected);
            Assert.Equal(CustomPropertyType.CLSID, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal(expected, prop.As<Guid>());
        }

        [Fact]
        public void CanCreateBasicBoolCustomProperty()
        {
            var prop = new CustomProperty("test", true);
            Assert.Equal(CustomPropertyType.Boolean, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.True(prop.As<bool>());
        }

        [Fact]
        public void CanCreateBasicDoubleCustomProperty()
        {
            double expected = 250.3221234;
            var prop = new CustomProperty("test", expected);
            Assert.Equal(CustomPropertyType.R8, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal(expected, prop.As<double>());
        }

        [Fact]
        public void CanCreateBasicDecimalCustomProperty()
        {
            decimal expected = new(250.32);
            var prop = new CustomProperty("test", expected);
            Assert.Equal(CustomPropertyType.Decimal, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal(expected, prop.As<decimal>());
        }

        [Fact]
        public void CanCreateBasicUintCustomProperty()
        {
            uint expected = uint.MaxValue;
            var prop = new CustomProperty("test", expected);
            Assert.Equal(CustomPropertyType.UnsignedInteger, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal(expected, prop.As<uint>());
        }

        [Fact]
        public void CanCreateCurrencyCustomProperty()
        {
            decimal expected = new(123.45);
            var prop = new CustomProperty("test", CustomPropertyType.Currency, expected);
            Assert.Equal(CustomPropertyType.Currency, prop.Type);
            Assert.Equal("test", prop.Name);
            Assert.Equal(expected, prop.As<decimal>());
            Assert.Equal((double)expected, prop.As<double>());
        }

        [Fact]
        public void CanCreateCustomPropertyFromXml()
        {
            var xe = XElement.Parse(
                @"<property name=""MyProperty"" pid=""2"" fmtid=""{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"" 
                    xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"" xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"">
                   <vt:i4>12345</vt:i4>
                 </property>");

            var prop = new CustomProperty(xe);
            Assert.Equal(CustomPropertyType.I4, prop.Type);
            Assert.Equal(12345, prop.As<int>());
        }

        [Fact]
        public void CtorChecksNamespace()
        {
            var xe = XElement.Parse(
                @"<property name=""MyProperty"" pid=""2"" fmtid=""{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"">
                 </property>");
            Assert.Throws<ArgumentException>(() => new CustomProperty(xe));
        }

        [Fact]
        public void CanChangeValue()
        {
            var expected = 2021;
            var prop = new CustomProperty("test", expected);
            Assert.Equal(CustomPropertyType.Integer, prop.Type);
            Assert.Equal(expected, prop.As<int>());

            prop.Value = "2022";
            Assert.Equal(2022, prop.As<int>());
        }

        [Fact]
        public void SetValueInteger()
        {
            var prop = new CustomProperty("test", 10);
            prop.SetValue(20);
            Assert.Equal(20, prop.As<int>());
        }

        [Fact]
        public void SetValueDateTime()
        {
            var prop = new CustomProperty("test", new DateTime(2020, 12, 25));
            Assert.Throws<ArgumentOutOfRangeException>(() => prop.SetValue(20));

            var dt = new DateTime(2022, 4, 1);
            prop.SetValue(dt);
            Assert.Equal(dt, prop.As<DateTime>());
        }

        [Fact]
        public void SetValueConvertsNumerics()
        {
            var prop = new CustomProperty("test", 10);
            prop.SetValue(20);
            Assert.Equal(20, prop.As<int>());

            prop.SetValue(21.23);
            Assert.Equal(21, prop.As<int>());

            prop.SetValue("15");
            Assert.Equal(15, prop.As<int>());

            prop.SetValue(new decimal(12345.3231255));
            Assert.Equal(12345, prop.As<int>());
        }

        [Fact]
        public void SetValueHandlesErrorCodeType()
        {
            string text = "0x8002042";
            var value = Convert.ToUInt32(text, 16);
            var prop = new CustomProperty("test", CustomPropertyType.ErrorCode, text);
            Assert.Equal(text, prop.Value);

            prop = new CustomProperty("test", CustomPropertyType.ErrorCode, value);
            Assert.Equal(text, prop.Value);
        }

        [Fact]
        public void SetValueThrowsOnOverflow()
        {
            var prop = new CustomProperty("test", CustomPropertyType.I1, 25);
            Assert.Equal(CustomPropertyType.I1, prop.Type);

            Assert.Throws<OverflowException>(() => prop.SetValue(128));

            sbyte val = 26;
            prop.SetValue(val);
            Assert.Equal(val, prop.As<sbyte>());
        }

        [Fact]
        public void SetValueThrowsOnInvalidType()
        {
            var prop = new CustomProperty("test", Guid.NewGuid());
            Assert.Throws<ArgumentOutOfRangeException>(()=>prop.SetValue(20));
        }
    }
}
