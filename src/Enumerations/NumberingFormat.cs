using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Specifies the numbering format which shall be used for a group of automatically numbered objects
/// </summary>
public enum NumberingFormat
{
    /// <summary>
    /// Bullet
    /// </summary>
    [XmlAttribute("bullet")] Bullet,

    /// <summary>
    /// Numbered
    /// </summary>
    [XmlAttribute("decimal")] Numbered,

    /// <summary>
    /// Upper roman
    /// </summary>
    UpperRoman,

    /// <summary>
    /// Lower roman
    /// </summary>
    LowerRoman,

    /// <summary>
    /// Uppercase letter
    /// </summary>
    UpperLetter,

    /// <summary>
    /// Lowercase letter
    /// </summary>
    LowerLetter,

    /// <summary>
    /// Ordinal
    /// </summary>
    Ordinal,

    /// <summary>
    /// Cardinal text
    /// </summary>
    CardinalText,

    /// <summary>
    /// Ordinal text
    /// </summary>
    OrdinalText,

    /// <summary>
    /// Hex values
    /// </summary>
    Hex,

    /// <summary>
    /// Chicago
    /// </summary>
    Chicago,

    /// <summary>
    /// Ideograph
    /// </summary>
    IdeographDigital,

    /// <summary>
    /// Japanese
    /// </summary>
    JapaneseCounting,

    /// <summary>
    /// Aiueo
    /// </summary>
    Aiueo,

    /// <summary>
    /// Iroha
    /// </summary>
    Iroha,

    /// <summary>
    /// Decimal - full width
    /// </summary>
    DecimalFullWidth,

    /// <summary>
    /// Decimal - half width
    /// </summary>
    DecimalHalfWidth,

    /// <summary>
    /// Japanese legal
    /// </summary>
    JapaneseLegal,

    /// <summary>
    /// Japanese digital
    /// </summary>
    JapaneseDigitalTenThousand,

    /// <summary>
    /// Digits in circles
    /// </summary>
    DecimalEnclosedCircle,

    /// <summary>
    /// Decimal full width
    /// </summary>
    DecimalFullWidth2,

    /// <summary>
    /// Aiueo full width
    /// </summary>
    AiueoFullWidth,

    /// <summary>
    /// Iroha full width
    /// </summary>
    IrohaFullWidth,

    /// <summary>
    /// Decimal starting at zero
    /// </summary>
    DecimalZero,

    /// <summary>
    /// Ganada
    /// </summary>
    Ganada,

    /// <summary>
    /// Chosung
    /// </summary>
    Chosung,

    /// <summary>
    /// Decimal enclosed
    /// </summary>
    DecimalEnclosedFullstop,

    /// <summary>
    /// Decimal in parenthesis
    /// </summary>
    DecimalEnclosedParen,

    /// <summary>
    /// Decimal in circle (Chinese)
    /// </summary>
    DecimalEnclosedCircleChinese,

    /// <summary>
    /// Ideograph in circle
    /// </summary>
    IdeographEnclosedCircle,

    /// <summary>
    /// Traditional ideograph
    /// </summary>
    IdeographTraditional,

    /// <summary>
    /// Zodiac ideograph
    /// </summary>
    IdeographZodiac,

    /// <summary>
    /// Zodiac traditional ideograph
    /// </summary>
    IdeographZodiacTraditional,

    /// <summary>
    /// Taiwan
    /// </summary>
    TaiwaneseCounting,

    /// <summary>
    /// Ideograph - Taiwan
    /// </summary>
    IdeographLegalTraditional,

    /// <summary>
    /// Taiwan counting
    /// </summary>
    TaiwaneseCountingThousand,

    /// <summary>
    /// Taiwan digital
    /// </summary>
    TaiwaneseDigital,

    /// <summary>
    /// Chinese counting
    /// </summary>
    ChineseCounting,

    /// <summary>
    /// Chinese legal (simplified)
    /// </summary>
    ChineseLegalSimplified,

    /// <summary>
    /// Chinese counting
    /// </summary>
    ChineseCountingThousand,

    /// <summary>
    /// Korean
    /// </summary>
    KoreanDigital,

    /// <summary>
    /// Korean counting
    /// </summary>
    KoreanCounting,

    /// <summary>
    /// Korean legal
    /// </summary>
    KoreanLegal,

    /// <summary>
    /// Korean digital (2)
    /// </summary>
    KoreanDigital2,

    /// <summary>
    /// Vietnamese
    /// </summary>
    VietnameseCounting,

    /// <summary>
    /// Russian lowercase
    /// </summary>
    RussianLower,

    /// <summary>
    /// Russian uppercase
    /// </summary>
    RussianUpper,

    /// <summary>
    /// None
    /// </summary>
    None,

    /// <summary>
    /// Numbers with dashes
    /// </summary>
    NumberInDash,

    /// <summary>
    /// Hebrew (1)
    /// </summary>
    Hebrew1,

    /// <summary>
    /// Hebrew (2)
    /// </summary>
    Hebrew2,

    /// <summary>
    /// Arabic alphabetic
    /// </summary>
    ArabicAlpha,

    /// <summary>
    /// Arabic
    /// </summary>
    ArabicAbjad,

    /// <summary>
    /// Hindi w/ vowels
    /// </summary>
    HindiVowels,
    
    /// <summary>
    /// Hindi w/ consonants
    /// </summary>
    HindiConsonants,

    /// <summary>
    /// Hindi w/ numbers
    /// </summary>
    HindiNumbers,

    /// <summary>
    /// Hindi
    /// </summary>
    HindiCounting,

    /// <summary>
    /// Thai w/ letters
    /// </summary>
    ThaiLetters,

    /// <summary>
    /// Thai w/ numbers
    /// </summary>
    ThaiNumbers,

    /// <summary>
    /// Thai
    /// </summary>
    ThaiCounting,

    /// <summary>
    /// Paragraph has numbering style deliberately removed
    /// </summary>
    Removed
}