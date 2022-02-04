using System.Xml.Serialization;

namespace DXPlus
{
    /// <summary>
    /// Specifies the numbering format which shall be used for a group of automatically numbered objects
    /// </summary>
    public enum NumberingFormat
    {
        [XmlAttribute("bullet")]
        Bullet,
        [XmlAttribute("decimal")]
        Numbered,
        UpperRoman,
        LowerRoman,
        UpperLetter,
        LowerLetter,
        Ordinal,
        CardinalText,
        OrdinalText,
        Hex,
        Chicago,
        IdeographDigital,
        JapaneseCounting,
        Aiueo,
        Iroha,
        DecimalFullWidth,
        DecimalHalfWidth,
        JapaneseLegal,
        JapaneseDigitalTenThousand,
        DecimalEnclosedCircle,
        DecimalFullWidth2,
        AiueoFullWidth,
        IrohaFullWidth,
        DecimalZero,
        Ganada,
        Chosung,
        DecimalEnclosedFullstop,
        DecimalEnclosedParen,
        DecimalEnclosedCircleChinese,
        IdeographEnclosedCircle,
        IdeographTraditional,
        IdeographZodiac,
        IdeographZodiacTraditional,
        TaiwaneseCounting,
        IdeographLegalTraditional,
        TaiwaneseCountingThousand,
        TaiwaneseDigital,
        ChineseCounting,
        ChineseLegalSimplified,
        ChineseCountingThousand,
        KoreanDigital,
        KoreanCounting,
        KoreanLegal,
        KoreanDigital2,
        VietnameseCounting,
        RussianLower,
        RussianUpper,
        None,
        NumberInDash,
        Hebrew1,
        Hebrew2,
        ArabicAlpha,
        ArabicAbjad,
        HindiVowels,
        HindiConsonants,
        HindiNumbers,
        HindiCounting,
        ThaiLetters,
        ThaiNumbers,
        ThaiCounting,
        Removed
    }
}