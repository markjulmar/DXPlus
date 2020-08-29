﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class RunTests
    {
        [Fact]
        public void SplitRunReturnsBothSides()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitRun(10);

            Assert.Equal("This ", results[0].Value);
            Assert.Equal("is a test.", results[1].Value);
        }

        [Fact]
        public void SplitAtZeroRunReturnsRightSide()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitRun(5);

            Assert.Null(results[0]);
            Assert.Equal("This is a test.", results[1].Value);
        }

        [Fact]
        public void SplitAtLengthRunReturnsLeftSide()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitRun(5 + text.Length);

            Assert.Null(results[1]);
            Assert.Equal("This is a test.", results[0].Value);
        }

    }
}
