from roman import toRoman
from numbers_parser.generated import TSWPArchives_pb2 as TSWPArchives


BULLET_PREFIXES = {
    TSWPArchives.ListStyleArchive.kNumericDecimal: "",
    TSWPArchives.ListStyleArchive.kNumericDoubleParen: "(",
    TSWPArchives.ListStyleArchive.kNumericRightParen: "",
    TSWPArchives.ListStyleArchive.kRomanUpperDecimal: "",
    TSWPArchives.ListStyleArchive.kRomanUpperDoubleParen: "(",
    TSWPArchives.ListStyleArchive.kRomanUpperRightParen: "",
    TSWPArchives.ListStyleArchive.kRomanLowerDecimal: "",
    TSWPArchives.ListStyleArchive.kRomanLowerDoubleParen: "(",
    TSWPArchives.ListStyleArchive.kRomanLowerRightParen: "",
    TSWPArchives.ListStyleArchive.kAlphaUpperDecimal: "",
    TSWPArchives.ListStyleArchive.kAlphaUpperDoubleParen: "(",
    TSWPArchives.ListStyleArchive.kAlphaUpperRightParen: "",
    TSWPArchives.ListStyleArchive.kAlphaLowerDecimal: "",
    TSWPArchives.ListStyleArchive.kAlphaLowerDoubleParen: "(",
    TSWPArchives.ListStyleArchive.kAlphaLowerRightParen: "",
}

BULLET_CONVERTION = {
    TSWPArchives.ListStyleArchive.kNumericDecimal: lambda x: str(x),
    TSWPArchives.ListStyleArchive.kNumericDoubleParen: lambda x: str(x),
    TSWPArchives.ListStyleArchive.kNumericRightParen: lambda x: str(x),
    TSWPArchives.ListStyleArchive.kRomanUpperDecimal: lambda x: toRoman(x),
    TSWPArchives.ListStyleArchive.kRomanUpperDoubleParen: lambda x: toRoman(x),
    TSWPArchives.ListStyleArchive.kRomanUpperRightParen: lambda x: toRoman(x),
    TSWPArchives.ListStyleArchive.kRomanLowerDecimal: lambda x: toRoman(x),
    TSWPArchives.ListStyleArchive.kRomanLowerDoubleParen: lambda x: toRoman(x),
    TSWPArchives.ListStyleArchive.kRomanLowerRightParen: lambda x: toRoman(x),
    TSWPArchives.ListStyleArchive.kAlphaUpperDecimal: lambda x: chr(x + 65),
    TSWPArchives.ListStyleArchive.kAlphaUpperDoubleParen: lambda x: chr(x + 65),
    TSWPArchives.ListStyleArchive.kAlphaUpperRightParen: lambda x: chr(x + 65),
    TSWPArchives.ListStyleArchive.kAlphaLowerDecimal: lambda x: chr(x + 97),
    TSWPArchives.ListStyleArchive.kAlphaLowerDoubleParen: lambda x: chr(x + 97),
    TSWPArchives.ListStyleArchive.kAlphaLowerRightParen: lambda x: chr(x + 97),
}

BULLET_SUFFIXES = {
    TSWPArchives.ListStyleArchive.kNumericDecimal: "",
    TSWPArchives.ListStyleArchive.kNumericDoubleParen: ")",
    TSWPArchives.ListStyleArchive.kNumericRightParen: ")",
    TSWPArchives.ListStyleArchive.kRomanUpperDecimal: ".",
    TSWPArchives.ListStyleArchive.kRomanUpperDoubleParen: ")",
    TSWPArchives.ListStyleArchive.kRomanUpperRightParen: ")",
    TSWPArchives.ListStyleArchive.kRomanLowerDecimal: ".",
    TSWPArchives.ListStyleArchive.kRomanLowerDoubleParen: ")",
    TSWPArchives.ListStyleArchive.kRomanLowerRightParen: ")",
    TSWPArchives.ListStyleArchive.kAlphaUpperDecimal: ".",
    TSWPArchives.ListStyleArchive.kAlphaUpperDoubleParen: ")",
    TSWPArchives.ListStyleArchive.kAlphaUpperRightParen: ")",
    TSWPArchives.ListStyleArchive.kAlphaLowerDecimal: ".",
    TSWPArchives.ListStyleArchive.kAlphaLowerDoubleParen: ")",
    TSWPArchives.ListStyleArchive.kAlphaLowerRightParen: ")",
}
