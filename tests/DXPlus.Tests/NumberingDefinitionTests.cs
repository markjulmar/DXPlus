using System.Globalization;
using Xunit;

namespace DXPlus.Tests
{
    public class NumberingDefinitionTests
    {
        [Fact]
        public void CustomListStartsAsSingleLevel()
        {
            var doc = Document.Create();
            var ns = doc.NumberingStyles;

            var nd = ns.AddCustomDefinition("-");
            Assert.Equal(NumberingLevelType.Single, nd.Style.LevelType);
            Assert.Single(nd.Style.Levels);
        }

        [Fact]
        public void CanAddNewLevels()
        {
            var doc = Document.Create();
            var ns = doc.NumberingStyles;

            var nd = ns.AddCustomDefinition("-");
            var style = nd.Style;

            var level = new NumberingLevel(1, NumberingFormat.DecimalEnclosedParen);
            Assert.Equal(720*2, level.ParagraphFormatting.LeftIndent);

            style.Levels.Add(level);
            Assert.Equal(2, style.Levels.Count);
            Assert.Equal(level, style.Levels[1]);
        }

        [Fact]
        public void CanRemoveLevels()
        {
            var doc = Document.Create();
            var ns = doc.NumberingStyles;

            uint creatorId = 0x1f;
            ns.AddNumberingStyle(new NumberingStyle
            {
                CreatorId = creatorId,
                LevelType = NumberingLevelType.Single,
                Levels = { new NumberingLevel(0) },
                Name = "test"
            });

            Assert.Equal(creatorId, ns.NumberingStyles[0].CreatorId);

            var nd = ns.CreateNumberingDefinition(ns.NumberingStyles[0]);
            Assert.Equal(creatorId, nd.Style.CreatorId);
            Assert.True(nd.Style.Id != -1);
            Assert.True(nd.Id > 0);
        }
    }
}
